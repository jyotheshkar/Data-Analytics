import unittest
import pandas as pd
from FinalAnalytics import weekly_summary

class SummaryTests(unittest.TestCase):
    def test_constant_week_counts_have_finite_zero_zscore(self):
        result, mean, std, threshold = weekly_summary(pd.DataFrame({'Created Date':['02/11/2026','09/11/2026']}))
        self.assertEqual(result['z-score'].tolist(), [0.0, 0.0])
        self.assertEqual(std, 0)
    def test_missing_week_is_counted(self):
        result, *_ = weekly_summary(pd.DataFrame({'Created Date':['02/11/2026','16/11/2026']}))
        self.assertEqual(result['Number of Bookings'].tolist(), [1,0,1])
    def test_invalid_inputs_fail_explicitly(self):
        for frame in [pd.DataFrame(), pd.DataFrame({'Created Date':[]}), pd.DataFrame({'Created Date':[None]})]:
            with self.assertRaises(ValueError): weekly_summary(frame)
    def test_does_not_mutate_input(self):
        frame = pd.DataFrame({'Created Date':['02/11/2026']})
        weekly_summary(frame)
        self.assertEqual(frame.columns.tolist(), ['Created Date'])
if __name__ == '__main__': unittest.main()
