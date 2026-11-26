"""Offline tests for the iCalendar / vCard conversions. Run: python3 -m unittest discover tests"""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
import davsync as d  # noqa: E402

TZ = ("BEGIN:VTIMEZONE\r\nTZID:GTB Standard Time\r\nX-ENTOURAGE-CFTIMEZONE:Europe/Bucharest\r\n"
      "BEGIN:STANDARD\r\nTZOFFSETFROM:+0300\r\nTZOFFSETTO:+0200\r\nDTSTART:16011001T040000\r\n"
      "END:STANDARD\r\nEND:VTIMEZONE\r\n")

OUTLOOK_ALLDAY = ("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Microsoft Corporation//Outlook for Mac MIMEDIR//EN\r\n"
                  "METHOD:PUBLISH\r\n" + TZ +
                  "BEGIN:VEVENT\r\nUID:outlook-uid\r\nX-ENTOURAGE_UUID:AAA\r\nX-MICROSOFT-EXCHANGE-ID:BBB\r\n"
                  "DTSTART;TZID=\"GTB Standard Time\":20261210T000000\r\n"
                  "DTEND;TZID=\"GTB Standard Time\":20261211T000000\r\nSUMMARY:All day\r\n"
                  "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n")

SERVER_EVENT = ("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//Example//EN\r\nBEGIN:VEVENT\r\nUID:srv-1\r\n"
                "DTSTAMP:20260101T000000Z\r\nDTSTART:20261201T080000Z\r\nDTEND:20261201T090000Z\r\n"
                "SUMMARY:Old title\r\nDESCRIPTION:Notes\r\nX-ALT-DESC;FMTTYPE=text/html:<p>Notes</p>\r\n"
                "ORGANIZER:mailto:boss@example.com\r\nATTENDEE;CN=Ana:mailto:ana@example.com\r\n"
                "BEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nTRIGGER:-PT15M\r\nEND:VALARM\r\n"
                "END:VEVENT\r\nEND:VCALENDAR\r\n")

OUTLOOK_EDIT = ("BEGIN:VCALENDAR\r\nVERSION:2.0\r\nMETHOD:PUBLISH\r\n" + TZ +
                "BEGIN:VEVENT\r\nUID:srv-1\r\nX-ENTOURAGE_UUID:AAA\r\n"
                "DTSTART;TZID=\"GTB Standard Time\":20261201T120000\r\n"
                "DTEND;TZID=\"GTB Standard Time\":20261201T130000\r\nSUMMARY:New title\r\n"
                "DESCRIPTION:Notes\r\nX-MICROSOFT-CDO-ALLDAYEVENT:FALSE\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n")

OUTLOOK_VCARD = ("begin:vcard\r\nversion:3.0\r\nprodid:Microsoft-MacOutlook/16\r\n"
                 "UID:BE000000-0000-0000-0000-000000000000\r\nfn;charset=utf-8:Ion Popescu\r\n"
                 "n;charset=utf-8:Popescu;Ion;;;\r\ntel;charset=utf-8;type=cell:+40700000000\r\n"
                 "X-MICROSOFT-EXCHANGE-ID:CCC\r\nend:vcard\r\n")


class Transforms(unittest.TestCase):
    def test_allday_becomes_date_and_outlook_ids_stripped(self):
        out = d.outlook_ics_to_server(OUTLOOK_ALLDAY, "srv-uid")
        self.assertIn("DTSTART;VALUE=DATE:20261210", out)
        self.assertIn("DTEND;VALUE=DATE:20261211", out)
        self.assertIn("UID:srv-uid", out)
        for gone in ("METHOD:", "X-ENTOURAGE_UUID", "X-MICROSOFT-EXCHANGE-ID"):
            self.assertNotIn(gone, out)
        self.assertIn("X-ENTOURAGE-CFTIMEZONE", out)  # timezone hint is kept

    def test_merge_keeps_server_only_fields(self):
        out = d.merge_ics(SERVER_EVENT, OUTLOOK_EDIT, "srv-1")
        self.assertIn("SUMMARY:New title", out)
        self.assertNotIn("Old title", out)
        self.assertIn("ATTENDEE;CN=Ana:mailto:ana@example.com", out)
        self.assertIn("ORGANIZER:mailto:boss@example.com", out)
        self.assertIn("BEGIN:VALARM", out)
        self.assertIn("X-ALT-DESC", out)  # description unchanged, so HTML copy stays
        self.assertIn("BEGIN:VTIMEZONE", out)  # Outlook's TZID needs its definition
        self.assertEqual(out.count("BEGIN:VEVENT"), 1)

    def test_merge_drops_stale_html_description(self):
        out = d.merge_ics(SERVER_EVENT, OUTLOOK_EDIT.replace("DESCRIPTION:Notes", "DESCRIPTION:Changed"), "srv-1")
        self.assertNotIn("X-ALT-DESC", out)

    def test_fingerprint_ignores_alarms_and_timestamps(self):
        a = SERVER_EVENT
        b = SERVER_EVENT.replace("DTSTAMP:20260101T000000Z", "DTSTAMP:20270101T000000Z")
        b = b.replace("BEGIN:VALARM\r\nACTION:DISPLAY\r\nDESCRIPTION:Reminder\r\nTRIGGER:-PT15M\r\nEND:VALARM\r\n", "")
        self.assertEqual(d.fingerprint(a), d.fingerprint(b))
        self.assertNotEqual(d.fingerprint(a), d.fingerprint(a.replace("Old title", "Other")))

    def test_vcard_normalised(self):
        out = d.outlook_vcf_to_server(OUTLOOK_VCARD, "card-1")
        self.assertTrue(out.startswith("BEGIN:VCARD\r\nVERSION:3.0"))
        self.assertIn("UID:card-1", out)
        self.assertIn("FN:Ion Popescu", out)
        self.assertIn("TEL;TYPE=cell:+40700000000", out)
        self.assertNotIn("charset", out.lower())
        self.assertNotIn("EXCHANGE", out)

    def test_vcard_merge_keeps_photo(self):
        server = ("BEGIN:VCARD\r\nVERSION:3.0\r\nUID:card-1\r\nFN:Old\r\nN:Old;;;;\r\n"
                  "PHOTO;ENCODING=b;TYPE=JPEG:AAAA\r\nEND:VCARD\r\n")
        out = d.merge_vcf(server, OUTLOOK_VCARD, "card-1")
        self.assertIn("PHOTO;ENCODING=b;TYPE=JPEG:AAAA", out)
        self.assertIn("FN:Ion Popescu", out)
        self.assertNotIn("FN:Old", out)
        self.assertEqual(out.count("UID:"), 1)

    def test_fold_respects_utf8_octets(self):
        line = "DESCRIPTION:" + "ă" * 60
        for part in d.fold(line).split("\r\n"):
            self.assertLessEqual(len(part.encode("utf-8")), 75)
        self.assertEqual(d.unfold(d.fold(line)), [line])


if __name__ == "__main__":
    unittest.main()
