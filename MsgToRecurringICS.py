#!/usr/bin/env python3
"""Convert multi-day Outlook MSG calendar invitations into weekday-only ICS series.

Uses only the Python standard library. Supported source time zone: Windows
'W. Europe Standard Time' (modern Central European DST rules, 1996 onwards).
No Outlook, COM, third-party package, or network access is required.

Authors: Ewald (GitHub repository owner) and ChatGPT (OpenAI).
Licence: MIT (see LICENSE).
"""
from __future__ import annotations

import argparse
import glob
import hashlib
import re
import struct
import sys
import uuid
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

VERSION = '1.0.0'
FREE = 0xFFFFFFFF
END = 0xFFFFFFFE
WINDOWS_TZ = 'W. Europe Standard Time'
MAX_FILE_SIZE = 200 * 1024 * 1024
# Numeric MAPI IDs are interpreted only in the Appointment property set.
APPOINTMENT_GUID = '00062002-0000-0000-c000-000000000046'
MEETING_GUID = '6ed8da90-450b-101b-98da-00aa003f1305'
PUBLIC_STRINGS_GUID = '00020329-0000-0000-c000-000000000046'


class ConversionError(Exception):
    """Input format, validation, or conversion error."""


class CompoundFile:
    """Minimal read-only Microsoft Compound File Binary Format (CFBF) reader."""

    def __init__(self, filename: Path):
        self.filename = filename
        self.raw = filename.read_bytes()
        if len(self.raw) > MAX_FILE_SIZE:
            raise ConversionError('MSG exceeds the 200 MiB safety limit')
        if self.raw[:8] != bytes.fromhex('D0CF11E0A1B11AE1'):
            raise ConversionError('Not a binary Outlook MSG (invalid CFBF signature)')
        if len(self.raw) < 512:
            raise ConversionError('Truncated compound-file header')
        h = self.raw[:512]
        self.sector_size = 1 << struct.unpack_from('<H', h, 30)[0]
        self.mini_sector_size = 1 << struct.unpack_from('<H', h, 32)[0]
        if self.sector_size not in (512, 4096) or self.mini_sector_size != 64:
            raise ConversionError('Unsupported CFBF sector sizes')
        self.cutoff = struct.unpack_from('<I', h, 56)[0]
        fat_count = struct.unpack_from('<I', h, 44)[0]
        dir_start = struct.unpack_from('<I', h, 48)[0]
        mini_start = struct.unpack_from('<I', h, 60)[0]
        mini_count = struct.unpack_from('<I', h, 64)[0]
        dif_start = struct.unpack_from('<I', h, 68)[0]
        dif_count = struct.unpack_from('<I', h, 72)[0]
        if fat_count > 200000 or dif_count > 200000 or mini_count > 200000:
            raise ConversionError('Implausible CFBF allocation-table sizes')

        difat = list(struct.unpack_from('<109I', h, 76))
        current = dif_start
        for _ in range(dif_count):
            if current in (FREE, END):
                raise ConversionError('Truncated DIFAT chain')
            sector = self._sector(current)
            difat += list(struct.unpack_from('<' + 'I' * (self.sector_size // 4 - 1), sector))
            current = struct.unpack_from('<I', sector, self.sector_size - 4)[0]
        fat_sectors = [i for i in difat if i != FREE][:fat_count]
        if len(fat_sectors) != fat_count:
            raise ConversionError('Incomplete FAT list')
        self.fat = []
        for sector_id in fat_sectors:
            self.fat.extend(struct.unpack('<' + 'I' * (self.sector_size // 4), self._sector(sector_id)))

        directory = self._chain(dir_start, self.fat, mini=False)
        self.entries = []
        for offset in range(0, len(directory) - 127, 128):
            entry = directory[offset:offset + 128]
            name_len = struct.unpack_from('<H', entry, 64)[0]
            if name_len > 64 or name_len % 2:
                raise ConversionError('Invalid CFBF directory name length')
            name = entry[:max(0, name_len - 2)].decode('utf-16le', errors='replace')
            self.entries.append(dict(
                name=name, kind=entry[66],
                left=struct.unpack_from('<I', entry, 68)[0],
                right=struct.unpack_from('<I', entry, 72)[0],
                child=struct.unpack_from('<I', entry, 76)[0],
                start=struct.unpack_from('<I', entry, 116)[0],
                length=struct.unpack_from('<Q', entry, 120)[0],
            ))
        if not self.entries or self.entries[0]['kind'] != 5:
            raise ConversionError('No CFBF root directory')
        self.minifat = []
        if mini_count and mini_start not in (FREE, END):
            mini_table = self._chain(mini_start, self.fat, mini=False)
            self.minifat = list(struct.unpack('<' + 'I' * (len(mini_table) // 4), mini_table))
        root = self.entries[0]
        self.mini_stream = self._chain(root['start'], self.fat, mini=False)[:root['length']]
        self.paths = {}
        self._visit(root['child'], '', set())

    def _sector(self, sid: int) -> bytes:
        start = (sid + 1) * self.sector_size
        end = start + self.sector_size
        if sid >= 0xFFFFFFFA or end > len(self.raw):
            raise ConversionError(f'Invalid CFBF sector reference: {sid}')
        return self.raw[start:end]

    def _chain(self, start: int, table: list[int], *, mini: bool) -> bytes:
        if start in (FREE, END):
            return b''
        out = bytearray()
        seen = set()
        current = start
        # All chains are bounded by the file's available sectors.
        while current not in (FREE, END):
            if current >= len(table) or current in seen:
                raise ConversionError('Invalid or cyclic CFBF allocation chain')
            seen.add(current)
            if mini:
                pos = current * self.mini_sector_size
                sector = self.mini_stream[pos:pos + self.mini_sector_size]
                if len(sector) != self.mini_sector_size:
                    raise ConversionError('Truncated CFBF mini stream')
            else:
                sector = self._sector(current)
            out.extend(sector)
            if len(out) > MAX_FILE_SIZE:
                raise ConversionError('Allocation chain exceeds safety limit')
            current = table[current]
        return bytes(out)

    def _visit(self, node_id: int, parent: str, visited: set[int]) -> None:
        if node_id == FREE:
            return
        if node_id >= len(self.entries) or node_id in visited:
            raise ConversionError('Invalid or cyclic CFBF directory tree')
        visited.add(node_id)
        node = self.entries[node_id]
        self._visit(node['left'], parent, visited) if node['left'] != FREE else None
        if node['kind'] in (1, 2):
            path = f"{parent}/{node['name']}" if parent else node['name']
            self.paths[path] = node_id
            if node['kind'] == 1 and node['child'] != FREE:
                self._visit(node['child'], path, set())
        self._visit(node['right'], parent, visited) if node['right'] != FREE else None

    def get(self, path: str) -> bytes | None:
        node_id = self.paths.get(path)
        if node_id is None:
            return None
        entry = self.entries[node_id]
        if entry['kind'] != 2:
            return None
        if entry['length'] > MAX_FILE_SIZE:
            raise ConversionError('Stream exceeds safety limit')
        content = self._chain(entry['start'], self.minifat if entry['length'] < self.cutoff else self.fat,
                              mini=entry['length'] < self.cutoff)
        if len(content) < entry['length']:
            raise ConversionError('Truncated CFBF stream: ' + path)
        return content[:entry['length']]


class OutlookMsg:
    """Read selected MAPI appointment properties and Outlook categories."""

    def __init__(self, path: Path):
        self.cfb = CompoundFile(path)
        self.properties = {}
        prop_data = self.cfb.get('__properties_version1.0')
        if prop_data is None or len(prop_data) < 32:
            raise ConversionError('Missing Outlook message property stream')
        for offset in range(32, len(prop_data) - 15, 16):
            tag = struct.unpack_from('<I', prop_data, offset)[0]
            self.properties[tag] = prop_data[offset + 8:offset + 16]

        self.named = {}
        root = '__nameid_version1.0/'
        entries = self.cfb.get(root + '__substg1.0_00030102')
        names = self.cfb.get(root + '__substg1.0_00040102')
        guid_data = self.cfb.get(root + '__substg1.0_00020102') or b''
        guid_map = {1: '00020328-0000-0000-c000-000000000046',
                    2: PUBLIC_STRINGS_GUID}
        for position in range(0, len(guid_data) - 15, 16):
            guid_map[3 + position // 16] = str(uuid.UUID(bytes_le=guid_data[position:position + 16]))
        if entries is None or names is None:
            raise ConversionError('Missing MSG named-property mapping')
        for offset in range(0, len(entries) - 7, 8):
            lid, kind_and_guid, index = struct.unpack_from('<IHH', entries, offset)
            guid_idx = kind_and_guid >> 1
            guid_key = guid_map.get(guid_idx)
            if not guid_key:
                continue
            if kind_and_guid & 1:
                if lid + 4 > len(names):
                    continue
                length = struct.unpack_from('<I', names, lid)[0]
                if length > len(names) - lid - 4:
                    continue
                key = names[lid + 4:lid + 4 + length].decode('utf-16le', 'replace')
                self.named[(guid_key, key)] = 0x8000 + index
            else:
                self.named[(guid_key, lid)] = 0x8000 + index

    def raw_property(self, prop_id: int, type_id: int) -> bytes | None:
        tag = (prop_id << 16) | type_id
        if type_id in (0x001F, 0x001E, 0x0102, 0x101F):
            return self.cfb.get(f'__substg1.0_{prop_id:04X}{type_id:04X}')
        return self.properties.get(tag)

    def string(self, prop_id: int) -> str:
        data = self.raw_property(prop_id, 0x001F)
        if data is not None:
            return data.decode('utf-16le', 'replace').rstrip('\x00')
        data = self.raw_property(prop_id, 0x001E)
        if data is not None:
            return data.decode('cp1252', 'replace').rstrip('\x00')
        return ''

    def named_id(self, guid: str, name: str | int) -> int | None:
        return self.named.get((guid, name))

    def named_bool(self, guid: str, lid: int) -> bool | None:
        prop_id = self.named_id(guid, lid)
        if prop_id is None:
            return None
        raw = self.raw_property(prop_id, 0x000B)
        return bool(struct.unpack_from('<H', raw)[0]) if raw else None

    def named_date(self, guid: str, lid: int) -> datetime | None:
        prop_id = self.named_id(guid, lid)
        if prop_id is None:
            return None
        return filetime_date(self.raw_property(prop_id, 0x0040))

    def categories(self) -> list[str]:
        # PidNameKeywords in PS_PUBLIC_STRINGS is PT_MV_UNICODE (0x101F).
        prop_id = self.named_id(PUBLIC_STRINGS_GUID, 'Keywords')
        if prop_id is None:
            return []
        prefix = f'__substg1.0_{prop_id:04X}101F-'
        indexed = []
        for path in self.cfb.paths:
            if path.startswith(prefix) and re.fullmatch(r'[0-9A-Fa-f]{8}', path[len(prefix):]):
                raw = self.cfb.get(path)
                if raw is not None:
                    indexed.append((int(path[len(prefix):], 16), raw.decode('utf-16le', 'replace').rstrip('\x00').strip()))
        return [s for _, s in sorted(indexed) if s]

    def timezone_key(self) -> str | None:
        # PidLidTimeZoneDefinitionStartDisplay (0x825E), MS-OXOCAL format:
        # header (8 bytes) includes a UTF-16LE Windows time-zone key length.
        for lid in (0x825E, 0x825F):
            prop_id = self.named_id(APPOINTMENT_GUID, lid)
            data = self.raw_property(prop_id, 0x0102) if prop_id is not None else None
            if data and len(data) >= 8:
                chars = struct.unpack_from('<H', data, 6)[0]
                if 0 < chars <= 255 and len(data) >= 8 + 2 * chars:
                    return data[8:8 + 2 * chars].decode('utf-16le', 'replace')
        return None

    def global_id(self) -> bytes | None:
        prop_id = self.named_id(MEETING_GUID, 0x0003)
        return self.raw_property(prop_id, 0x0102) if prop_id is not None else None


def filetime_date(raw: bytes | None) -> datetime | None:
    if not raw or len(raw) < 8:
        return None
    ticks = struct.unpack_from('<Q', raw)[0]
    try:
        return datetime(1601, 1, 1, tzinfo=timezone.utc) + timedelta(microseconds=ticks // 10)
    except OverflowError as exc:
        raise ConversionError('Invalid MSG FILETIME value') from exc


def last_sunday(year: int, month: int) -> date:
    if month == 12:
        last = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        last = date(year, month + 1, 1) - timedelta(days=1)
    return last - timedelta(days=(last.weekday() + 1) % 7)


def western_europe_local(utc_dt: datetime) -> datetime:
    """Convert UTC into W. Europe Standard Time without system tzdata."""
    if utc_dt.tzinfo is None:
        raise ValueError('UTC time must be aware')
    utc = utc_dt.astimezone(timezone.utc)
    year = utc.year
    if year < 1996:
        raise ConversionError('The built-in European DST rule supports events from 1996 onwards')
    begins = datetime.combine(last_sunday(year, 3), time(1), tzinfo=timezone.utc)
    ends = datetime.combine(last_sunday(year, 10), time(1), tzinfo=timezone.utc)
    offset = 2 if begins <= utc < ends else 1
    return (utc + timedelta(hours=offset)).replace(tzinfo=None)


def get_appointment(msg: OutlookMsg) -> tuple[str, date, date, str, list[str], bytes]:
    msg_class = msg.string(0x001A)
    if not msg_class.startswith('IPM.Appointment'):
        raise ConversionError('MSG is not an Outlook calendar appointment: ' + (msg_class or 'unknown class'))
    allday = msg.named_bool(APPOINTMENT_GUID, 0x8215)
    if allday is not True:
        raise ConversionError('Only all-day appointment invitations are supported')
    tz_key = msg.timezone_key()
    if tz_key != WINDOWS_TZ:
        raise ConversionError(f'Unsupported input time zone {tz_key!r}; expected {WINDOWS_TZ!r}')
    utc_start = msg.named_date(APPOINTMENT_GUID, 0x820D)
    utc_end = msg.named_date(APPOINTMENT_GUID, 0x820E)
    if utc_start is None or utc_end is None:
        utc_start = utc_start or filetime_date(msg.raw_property(0x0060, 0x0040))
        utc_end = utc_end or filetime_date(msg.raw_property(0x0061, 0x0040))
    if utc_start is None or utc_end is None or utc_end <= utc_start:
        raise ConversionError('Missing or inconsistent appointment start/end dates')
    local_start = western_europe_local(utc_start)
    local_end = western_europe_local(utc_end)
    if local_start.time() != time.min or local_end.time() != time.min:
        raise ConversionError('All-day event boundaries do not fall on local midnight')
    start = local_start.date()
    end_exclusive = local_end.date()
    if end_exclusive <= start:
        raise ConversionError('Invalid all-day appointment date range')
    subject = msg.string(0x0037).strip()
    if not subject:
        raise ConversionError('MSG has an empty subject')
    match = re.match(r'^\s*([A-Za-z][A-Za-z0-9]*-\d+)(?=\b|:|\s|$)', subject)
    base = match.group(1) if match else subject.split(':', 1)[0].strip()
    if not base:
        raise ConversionError('Cannot derive appointment subject')
    categories = msg.categories()
    # First name in the first Outlook category, as requested by sample mappings.
    suffix = categories[0].split()[0] if categories else ''
    output_subject = base + (f' ({suffix})' if suffix else '')
    description = msg.string(0x1000).rstrip()
    stable_seed = msg.global_id() or (subject + utc_start.isoformat() + utc_end.isoformat()).encode('utf-8')
    return output_subject, start, end_exclusive, description, categories, stable_seed


def workdays(start: date, end_exclusive: date) -> list[date]:
    days = []
    day = start
    while day < end_exclusive:
        if day.weekday() < 5:
            days.append(day)
        day += timedelta(days=1)
    return days


def escape_ics(value: str) -> str:
    return (value.replace('\\', '\\\\').replace('\r\n', '\n').replace('\r', '\n')
            .replace('\n', '\\n').replace(';', '\\;').replace(',', '\\,'))


def fold_ics_line(line: str) -> str:
    """Fold long iCalendar content lines at 75 UTF-8 octets."""
    result = []
    chunk = ''
    nbytes = 0
    for ch in line:
        width = len(ch.encode('utf-8'))
        if nbytes + width > 75:
            result.append(chunk)
            chunk = ' '
            nbytes = 1
        chunk += ch
        nbytes += width
    result.append(chunk)
    return '\r\n'.join(result)


def build_ics(subject: str, description: str, days: list[date], start_at: time,
              end_at: time, stable_seed: bytes) -> bytes:
    first = days[0]
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
    uid = hashlib.sha256(stable_seed).hexdigest()[:32] + '@msgtorecurringics.local'
    lines = [
        'BEGIN:VCALENDAR',
        'PRODID:-//MsgToRecurringICS//Version 1.0//EN',
        'VERSION:2.0',
        'CALSCALE:GREGORIAN',
        'METHOD:PUBLISH',
        'BEGIN:VTIMEZONE',
        f'TZID:{WINDOWS_TZ}',
        'BEGIN:STANDARD',
        'DTSTART:19701025T030000',
        'RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU',
        'TZOFFSETFROM:+0200',
        'TZOFFSETTO:+0100',
        'END:STANDARD',
        'BEGIN:DAYLIGHT',
        'DTSTART:19700329T020000',
        'RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU',
        'TZOFFSETFROM:+0100',
        'TZOFFSETTO:+0200',
        'END:DAYLIGHT',
        'END:VTIMEZONE',
        'BEGIN:VEVENT',
        'CLASS:PUBLIC',
        f'DTSTAMP:{stamp}',
        f'DTSTART;TZID="{WINDOWS_TZ}":{first:%Y%m%d}T{start_at:%H%M%S}',
        f'DTEND;TZID="{WINDOWS_TZ}":{first:%Y%m%d}T{end_at:%H%M%S}',
        f'RRULE:FREQ=DAILY;BYDAY=MO,TU,WE,TH,FR;COUNT={len(days)}',
        f'SUMMARY:{escape_ics(subject)}',
        f'DESCRIPTION:{escape_ics(description)}',
        'PRIORITY:5',
        'SEQUENCE:0',
        'STATUS:CONFIRMED',
        'TRANSP:OPAQUE',
        f'UID:{uid}',
        'END:VEVENT',
        'END:VCALENDAR',
    ]
    return ('\r\n'.join(fold_ics_line(line) for line in lines) + '\r\n').encode('utf-8')


def parse_clock(value: str) -> time:
    try:
        return datetime.strptime(value, '%H:%M').time()
    except ValueError as exc:
        raise argparse.ArgumentTypeError('Expected 24-hour HH:MM, e.g. 08:30') from exc


def safe_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', name).rstrip(' .') or 'appointment'


@dataclass
class Result:
    source: Path
    target: Path
    subject: str
    dates: list[date]
    categories: list[str]
    written: bool


def convert(source: Path, directory: Path | None, start_at: time, end_at: time,
            *, overwrite: bool, dry_run: bool) -> Result:
    msg = OutlookMsg(source)
    subject, start, until, description, categories, seed = get_appointment(msg)
    dates = workdays(start, until)
    if not dates:
        raise ConversionError('The event covers no weekdays')
    destination = (directory or source.parent) / (safe_filename(subject) + '.ics')
    if destination.exists() and not overwrite and not dry_run:
        raise ConversionError(f'Output already exists: {destination} (use --overwrite)')
    contents = build_ics(subject, description, dates, start_at, end_at, seed)
    if not dry_run:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(contents)
    return Result(source, destination, subject, dates, categories, not dry_run)


def source_files(arguments: list[str]) -> list[Path]:
    found = []
    used = set()
    for argument in arguments:
        matches = [Path(p) for p in glob.glob(argument)] if glob.has_magic(argument) else [Path(argument)]
        if not matches:
            raise ConversionError('No files match input pattern: ' + argument)
        for path in matches:
            files = sorted((p for p in path.iterdir() if p.is_file() and p.suffix.lower() == '.msg'),
                           key=lambda p: str(p).lower()) if path.is_dir() else [path]
            for file in files:
                if not file.exists() or not file.is_file() or file.suffix.lower() != '.msg':
                    raise ConversionError('Not an existing MSG file: ' + str(file))
                key = str(file.resolve()).lower()
                if key not in used:
                    found.append(file)
                    used.add(key)
    return found


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description='Convert Outlook all-day MSG invitations to recurring weekday ICS appointments.',
        epilog='Example: python MsgToRecurringICS.py *.msg --output-dir output')
    parser.add_argument('inputs', nargs='+', metavar='MSG', help='MSG files, glob patterns or folders')
    parser.add_argument('--start', type=parse_clock, default=time(8, 30), metavar='HH:MM',
                        help='Daily start time (default: 08:30)')
    parser.add_argument('--end', type=parse_clock, default=time(16, 30), metavar='HH:MM',
                        help='Daily end time (default: 16:30)')
    parser.add_argument('--output-dir', type=Path, help='Output folder (default: beside each MSG)')
    parser.add_argument('--overwrite', action='store_true', help='Replace existing ICS files')
    parser.add_argument('--dry-run', action='store_true', help='Report conversions without writing')
    parser.add_argument('--version', action='version', version=VERSION)
    args = parser.parse_args(argv)
    if args.end <= args.start:
        parser.error('--end must be later than --start on the same day')
    try:
        inputs = source_files(args.inputs)
    except (ConversionError, OSError) as exc:
        parser.error(str(exc))
    if not inputs:
        parser.error('No MSG files found')
    failures = 0
    for source in inputs:
        print(f'--- {source.stem}')
        try:
            result = convert(source, args.output_dir, args.start, args.end,
                             overwrite=args.overwrite, dry_run=args.dry_run)
            print(f"Subject:    {result.subject}")
            print(f"Categories: {', '.join(result.categories) if result.categories else '(none)'}")
            print(f"Dates:      {result.dates[0]:%Y-%m-%d} to {result.dates[-1]:%Y-%m-%d} "
                  f'({len(result.dates)} weekdays)')
            print(f"Hours:      {args.start:%H:%M}-{args.end:%H:%M}")
            print(f"Output:     {result.target}" + (' [dry run]' if args.dry_run else ''))
        except (ConversionError, OSError, UnicodeError, struct.error, ValueError) as exc:
            failures += 1
            print(f'ERROR:      {exc}', file=sys.stderr)
    if failures:
        print(f'Failed: {failures} of {len(inputs)} inputs', file=sys.stderr)
        return 1
    print(f'Converted: {len(inputs)} file(s)' + (' [dry run]' if args.dry_run else ''))
    return 0


if __name__ == '__main__':
    sys.exit(main())
