import unittest

import pandas as pd

from mor_parser import (
    DetectedDataset,
    KUB_PAGE1_FIELDS,
    KUB_PAGE2_FIELDS,
    KUB_PAGE3_FIELDS,
    combine_same_named_datasets,
    extract_known_page,
    is_kub_fourth_creek,
)


class FakePage:
    def __init__(self, words):
        self.words = words

    def extract_words(self, **_kwargs):
        return self.words


class FakeTextPage:
    def __init__(self, text):
        self.text = text

    def extract_text(self):
        return self.text


class FakePdf:
    def __init__(self, page_texts):
        self.pages = [FakeTextPage(text) for text in page_texts]


def word(text, center, top=100.0):
    return {
        "text": str(text),
        "x0": center - 1.0,
        "x1": center + 1.0,
        "top": top,
    }


class KnownKubParserTests(unittest.TestCase):
    def test_three_page_fingerprint_recognizes_monthly_text_variations(self):
        pdf = FakePdf([
            (
                "Report of Operation of Wastewater Treatment Plant\n"
                "Influent Flows, MGD\nSet Solids\nFinal Effluent Parameters"
            ),
            (
                "Secondary System Digested Sludge Digester Influent "
                "Digester No. 2 Digester No. 4 Digester No. 6"
            ),
            (
                "Cadmium Chromium Copper Nickel Zinc Silver Lead "
                "Filter Press Solids Disposal"
            ),
        ])

        self.assertTrue(is_kub_fourth_creek(pdf))

    def test_known_page_schemas_are_complete_and_unique(self):
        self.assertEqual(len(KUB_PAGE1_FIELDS), 33)
        self.assertEqual(len(KUB_PAGE2_FIELDS), 25)
        self.assertEqual(len(KUB_PAGE3_FIELDS), 19)

        for field_defs in (KUB_PAGE1_FIELDS, KUB_PAGE2_FIELDS, KUB_PAGE3_FIELDS):
            names = [name for name, _ in field_defs]
            self.assertEqual(len(names), len(set(names)))
            self.assertEqual(names[0], "Date")

    def test_page_one_parent_headers_trickle_down_without_fused_neighbors(self):
        names = [name for name, _ in KUB_PAGE1_FIELDS]

        self.assertEqual(
            names[2:5],
            [
                "Influent Flows, MGD Avg",
                "Influent Flows, MGD Max",
                "Influent Flows, MGD Min",
            ],
        )
        self.assertEqual(
            names[24:33],
            [
                "Influent Parameters Total N",
                "Influent Parameters Total P",
                "Final Effluent Parameters E-Coli",
                "Final Effluent Parameters Cl2 Resid.",
                "Final Effluent Parameters Lbs Cl2",
                "Final Effluent Parameters NH3-N Comp.",
                "Final Effluent Parameters Total N",
                "Final Effluent Parameters Total P",
                "Final Effluent Parameters DO",
            ],
        )

        forbidden_fused_headers = {
            "Influent Total P E-Coli",
            "Final Effluent Parameters NH3-N Comp. Total N",
            "Total P DO",
        }
        self.assertTrue(forbidden_fused_headers.isdisjoint(names))

    def test_page_three_keeps_each_metal_in_its_own_column(self):
        values = {
            "Cadmium Influent mg/L": "0.0020",
            "Cadmium Effluent mg/L": "0.0020",
            "Chromium Influent mg/L": "0.0030",
            "Chromium Effluent mg/L": "0.0020",
            "Copper Influent mg/L": "0.0330",
            "Copper Effluent mg/L": "0.0080",
            "Nickel Influent mg/L": "0.0050",
            "Nickel Effluent mg/L": "0.0050",
            "Zinc Influent mg/L": "0.1260",
            "Zinc Effluent mg/L": "0.0430",
            "Silver Influent mg/L": "0.0020",
            "Silver Effluent mg/L": "0.0020",
            "Lead Influent mg/L": "0.0050",
            "Lead Effluent mg/L": "0.0050",
            "Filter Press % Solids": "24",
            "Solids Disposal Landfill Loads": "0",
        }
        anchors = dict(KUB_PAGE3_FIELDS)
        words = [word("01/06/2021", anchors["Date"])]
        words.extend(word(value, anchors[name]) for name, value in values.items())

        frame = extract_known_page(FakePage(words), KUB_PAGE3_FIELDS)

        self.assertEqual(frame.shape, (1, 19))
        self.assertEqual(list(frame.columns), [name for name, _ in KUB_PAGE3_FIELDS])
        for name, value in values.items():
            self.assertEqual(frame.loc[0, name], float(value))
        self.assertIsNone(frame.loc[0, "Filter Press Lbs of Solids"])
        self.assertIsNone(frame.loc[0, "Solids Disposal Farm Loads"])

    def test_monthly_page_three_datasets_combine_without_parallel_headers(self):
        columns = [name for name, _ in KUB_PAGE3_FIELDS]
        january = pd.DataFrame(
            [[pd.Timestamp("2021-01-06")] + [None] * (len(columns) - 1)],
            columns=columns,
        )
        march = pd.DataFrame(
            [[pd.Timestamp("2021-03-02")] + [None] * (len(columns) - 1)],
            columns=columns,
        )
        datasets = [
            DetectedDataset("PDF Page 3", "January.pdf", january, "High", []),
            DetectedDataset("PDF Page 3", "March.pdf", march, "High", []),
        ]

        combined = combine_same_named_datasets(datasets)

        self.assertEqual(len(combined), 1)
        self.assertEqual(combined[0].dataframe.shape, (2, 19))
        self.assertEqual(list(combined[0].dataframe.columns), columns)


if __name__ == "__main__":
    unittest.main()
