"""Run with: python -m unittest -v test_food_detection.py"""

import importlib.util
import sys
import tempfile
import types
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

# Allow the text-only scraper tests to run without installing its HTML/HTTP
# dependencies. GitHub Actions already installs the real dependencies.
if importlib.util.find_spec("bs4") is None:
    bs4 = types.ModuleType("bs4")
    bs4.BeautifulSoup = lambda *args, **kwargs: None
    sys.modules["bs4"] = bs4
if importlib.util.find_spec("requests") is None:
    requests = types.ModuleType("requests")
    requests.Response = type("Response", (), {})
    sys.modules["requests"] = requests

import harvard_food_events as scraper
import generate_calendar as calendar


TITLE = ("Visiting Scholar Presentation featuring Li Zhang — Defying Slow Death: "
         "Unsafe Food, Governance, and Quiet Social Movements in China")
URL = "https://fairbank.fas.harvard.edu/events/visiting-scholar-presentation-featuring-li-zhang-defying-slow-death-unsafe-food-governance-and-quiet-social-movements-in-china/"


class Element:
    def __init__(self, text="", attrs=None, children=None):
        self.text = text
        self.attrs = attrs or {}
        self.children = children or {}

    def get_text(self, *args, **kwargs):
        return self.text

    def get(self, key, default=""):
        return self.attrs.get(key, default)

    def select_one(self, selector):
        return self.children.get(selector)

    def select(self, selector):
        return self.children.get(selector, [])


def article(title, url):
    return Element(children={
        ".tribe-events-calendar-list__event-title a": Element(title, {"href": url}),
        "time[datetime]": Element("November 16 @ 3:30 pm – 4:30 pm",
                                  {"datetime": "2026-11-16"}),
        ".tribe-events-calendar-list__event-venue-title": Element("CGIS South"),
        "[class*='description']": Element("Speaker: Li Zhang. Food safety research."),
    })


class FoodDetectionTests(unittest.TestCase):
    def test_topic_words_are_not_food_offers(self):
        for text in (TITLE + " Speaker: Li Zhang.",
                     "Research on food safety, tea production, and lunch policy.",
                     "A talk about coffee and free food networks."):
            with self.subTest(text=text):
                self.assertFalse(scraper._has_food(text))
                self.assertEqual(scraper._food_snippet(text), "")

    def test_clear_offers_are_retained(self):
        for text in ("Light refreshments will be served after the lecture.",
                     "Lunch will be provided.",
                     "Please join us for lunch after the talk.",
                     "Coffee and tea are available."):
            with self.subTest(text=text):
                self.assertTrue(scraper._has_food(text))
                self.assertEqual(scraper._food_snippet(text), text.rstrip("."))

        long_text = TITLE + " Speaker: Li Zhang. Lunch will be provided."
        self.assertEqual(scraper._food_snippet(long_text), "Lunch will be provided")

    def test_fairbank_scrape_and_both_outputs(self):
        positive_url = "https://fairbank.fas.harvard.edu/events/lunch-seminar/"
        listing = Element(children={
            "article.tribe-events-calendar-list__event": [
                article(TITLE, URL), article("Lunch seminar", positive_url)
            ]
        })
        detail = Element("Lecture abstract about food safety. Lunch will be provided.")
        responses = {"https://fairbank.fas.harvard.edu/events/": "listing",
                     URL: "topic", positive_url: "detail"}

        with patch.object(scraper, "_get", side_effect=lambda url: types.SimpleNamespace(text=responses[url])), \
             patch.object(scraper, "BeautifulSoup", side_effect=lambda html, parser: listing if html == "listing" else detail if html == "detail" else Element(TITLE + " Speaker: Li Zhang. The talk examines food safety.")), \
             patch.object(scraper.time, "sleep"):
            start = datetime(2026, 11, 15, tzinfo=scraper.BOSTON_TZ)
            end = datetime(2026, 11, 22, tzinfo=scraper.BOSTON_TZ)
            events = scraper.fetch_tribe(scraper.HTML_CALENDARS[1], start, end)

        self.assertEqual([e["title"] for e in events], ["Lunch seminar"])
        self.assertEqual(events[0]["food_note"], "Lunch will be provided")

        with tempfile.TemporaryDirectory() as directory:
            readme = Path(directory) / "README.md"
            scraper.write_readme(events, str(readme), start)
            readme_text = readme.read_text(encoding="utf-8")
            ics = calendar.generate_ics(calendar.extract_events_from_readme(str(readme)))
            self.assertIn("Lunch seminar", readme_text)
            self.assertNotIn("Li Zhang", readme_text)
            self.assertIn("Lunch seminar", ics)
            self.assertNotIn("Li Zhang", ics)


if __name__ == "__main__":
    unittest.main()
