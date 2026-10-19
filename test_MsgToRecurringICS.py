#!/usr/bin/env python3
"""Unittest suite; sample pair integration tests use MSG/ICS files from ../."""
from __future__ import annotations

import os
import re
import tempfile
import unittest
from datetime import date, datetime, time, timezone
from pathlib import Path
from unittest.mock import patch

import MsgToRecurringICS as converter

ROOT = Path(__file__).resolve().parent
FIXTURES = Path(os.environ.get("MSG_TEST_FIXTURES", str(ROOT / "fixtures")))


def unfold(data: bytes) -> dict[str, str]:
    text = data.decode('utf-8')
    assert '\r\n' in text
    lines = re.sub(r'\r\n[ \t]', '', text).split('\r\n')
    out = {}
    for line in lines:
        if ':' in line:
            k, v = line.split(':', 1)
            out[k] = v
    return out


class ConverterTests(unittest.TestCase):
    def test_weekends_are_excluded(self):
        # Friday through Tuesday inclusive => Fri, Mon, Tue (three occurrences).
        days = converter.workdays(date(2027, 4, 16), date(2027, 4, 21))
        self.assertEqual(days, [date(2027, 4, 16), date(2027, 4, 19), date(2027, 4, 20)])
        ics = converter.build_ics('TRAIN-999 (Ewald)', 'Description', days,
                                  time(8, 30), time(16, 30), b'msg-identity')
        self.assertIn(b'BYDAY=MO,TU,WE,TH,FR;COUNT=3', ics)
        self.assertEqual(unfold(ics)['DTSTART;TZID="W. Europe Standard Time"'], '20270416T083000')

    def test_no_category_means_no_suffix(self):
        original = next(FIXTURES.glob('TRAIN-761_*.msg'), None)
        if original is None:
            self.skipTest('Set MSG_TEST_FIXTURES to a directory containing the original MSG files')
        msg = converter.OutlookMsg(original)
        with patch.object(msg, 'categories', return_value=[]):
            subject, *_ = converter.get_appointment(msg)
        self.assertEqual(subject, 'TRAIN-761')

    def test_custom_hours(self):
        days = converter.workdays(date(2027, 5, 10), date(2027, 5, 14))
        output = unfold(converter.build_ics('TRAIN-761', 'Details', days,
                                            time(9, 15), time(17), b'source'))
        self.assertEqual(output['DTSTART;TZID="W. Europe Standard Time"'], '20270510T091500')
        self.assertEqual(output['DTEND;TZID="W. Europe Standard Time"'], '20270510T170000')

    def test_dst_local_dates(self):
        winter = datetime(2027, 1, 5, 23, tzinfo=timezone.utc)
        summer = datetime(2027, 4, 12, 22, tzinfo=timezone.utc)
        self.assertEqual(converter.western_europe_local(winter).date(), date(2027, 1, 6))
        self.assertEqual(converter.western_europe_local(summer).date(), date(2027, 4, 13))
        self.assertEqual(converter.western_europe_local(summer).hour, 0)

    def test_ics_escaping_and_folding(self):
        val = 'Ö' * 90 + ',;\\' + '\nNext'
        content = converter.fold_ics_line('DESCRIPTION:' + converter.escape_ics(val))
        lines = content.split('\r\n')
        self.assertTrue(all(len(x.encode('utf-8')) <= 75 for x in lines))
        self.assertIn('\\nNext', content)
        self.assertIn('\\,\\;\\\\', content)

    def test_stable_uid(self):
        days = [date(2027, 5, 10)]
        a = unfold(converter.build_ics('A', '', days, time(8, 30), time(16, 30), b'constant'))
        b = unfold(converter.build_ics('B', '', days, time(9, 30), time(17, 30), b'constant'))
        self.assertEqual(a['UID'], b['UID'])

    def test_dry_run_can_target_existing_file(self):
        msg = next(FIXTURES.glob('TRAIN-761_*.msg'), None)
        if msg is None:
            self.skipTest('Set MSG_TEST_FIXTURES to a directory containing the original MSG files')
        with tempfile.TemporaryDirectory() as d:
            first = converter.convert(msg, Path(d), time(8, 30), time(16, 30), overwrite=False, dry_run=False)
            snapshot = first.target.read_bytes()
            dry = converter.convert(msg, Path(d), time(9), time(17), overwrite=False, dry_run=True)
            self.assertFalse(dry.written)
            self.assertEqual(first.target.read_bytes(), snapshot)
            with self.assertRaises(converter.ConversionError):
                converter.convert(msg, Path(d), time(9), time(17), overwrite=False, dry_run=False)


class ReferencePairTests(unittest.TestCase):
    def test_three_real_msg_and_ics_pairs(self):
        expected = {
            'TRAIN-761': ('20270510T083000', 4),
            'TRAIN-766': ('20270413T083000', 3),
            'TRAIN-770': ('20270420T083000', 3),
        }
        for train_id, (when, count) in expected.items():
            with self.subTest(training=train_id):
                source_msg = next(FIXTURES.glob(train_id + '_*.msg'), None)
                reference_ics = FIXTURES / (train_id + ' (Ewald).ics')
                if source_msg is None or not reference_ics.is_file():
                    self.skipTest('Set MSG_TEST_FIXTURES to a directory containing the three original MSG/ICS pairs')
                msg = converter.OutlookMsg(source_msg)
                self.assertEqual(msg.categories(), ['Ewald Müller'])
                subject, start, end, description, categories, seed = converter.get_appointment(msg)
                self.assertEqual(subject, train_id + ' (Ewald)')
                days = converter.workdays(start, end)
                self.assertEqual(len(days), count)
                output = unfold(converter.build_ics(subject, description, days,
                                                      time(8, 30), time(16, 30), seed))
                reference = unfold(reference_ics.read_bytes())
                self.assertEqual(output['SUMMARY'], reference['SUMMARY;LANGUAGE=en-at'])
                self.assertEqual(output['DTSTART;TZID="W. Europe Standard Time"'], when)
                self.assertEqual(output['DTSTART;TZID="W. Europe Standard Time"'],
                                 reference['DTSTART;TZID="W. Europe Standard Time"'])
                self.assertEqual(output['DTEND;TZID="W. Europe Standard Time"'],
                                 reference['DTEND;TZID="W. Europe Standard Time"'])
                self.assertIn('FREQ=DAILY', reference['RRULE'])
                self.assertEqual(int(reference['RRULE'].split('COUNT=')[1]), count)
                self.assertEqual(int(output['RRULE'].split('COUNT=')[1]), count)
                self.assertEqual(output['DESCRIPTION'].split('\\n')[0].strip(),
                                 reference['DESCRIPTION'].split('\\n')[0].strip())
                self.assertEqual(output['CLASS'], 'PUBLIC')
                self.assertEqual(output['TRANSP'], 'OPAQUE')
                self.assertNotIn('LOCATION', output)
                self.assertNotIn('BEGIN:VALARM', output)


if __name__ == '__main__':
    unittest.main(verbosity=2)
