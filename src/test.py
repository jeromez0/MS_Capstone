import unittest
import config
import data_analytics
from data_analytics import Analyticsdashboard

### Unit testing for the analytics. Makes sure everything is put together cleanly.

class Test(unittest.TestCase):

    def test_caseVOE(self): ### test case for VOE drafter feature
        #### these commands will initialize the GatherData class from data_analytics.py and prep the unit test.
        config.fields = 'C03:R1000'
        data_analytics.mainloopVOE()
        ### the main test that is being run will ensure that every list within the self.information primary list
        ### has 9 values within it.
        for i in config.information:
            length = len(i)
            self.assertEqual(length, 8)

    def test_caseJDReqs(self): ### testcase for JDReqs feature
        #### these commands will initialize the GatherData from data_analytics.py and prep the unit test.
        config.fields = 'C03:Q1000'
        data_analytics.mainloopJDReqs()

        ### the main test that is being run will ensure that every list within the self.information primary list
        ### has 9 values within it.
        for i in config.information:
            length = len(i)
            self.assertEqual(length, 10)

    def test_caseETAReview(self): ### testcase for ETAReview feature
        #### these commands will initialize the GatherData class from data_analytics.py and prep the unit test.
        config.fields = 'C03:AA1000'
        data_analytics.mainloopETAReview()

        ### the main test that is being run will ensure that every list within the self.information primary list
        ### has 9 values within it.
        for i in config.information:
            length = len(i)
            self.assertEqual(length, 6)

    def test_casePERMFiled(self): ### testcase for PERM Filed feature
        #### these commands will initialize the Gatherdata class from data_analytics.py and prep the unit test.
        config.fields = 'C03:AG1000'
        data_analytics.mainloopPERMFiled()

        ### the main test that is being run will ensure that every list within the self.information primary list
        ### has 9 values within it.
        for i in config.information:
            length = len(i)
            self.assertEqual(length,8)

    def test_casedashboard(self): ### testcase for dashboard feature
        #### this commands will initialize the analyticsdashboard from data_analytics.py and prep the unit test.
        analyze = Analyticsdashboard()
        analyze.gatherthedata()

        #### these tests will ensure that all the lists are the proper length, all the mathematics is correct
        self.assertEqual(len(analyze.figurespie), 3, len(analyze.labelspie))
        self.assertEqual(len(analyze.figuresbar), 2, len(analyze.labelsbar))
        self.assertEqual(analyze.counter, analyze.confirmed['Confirmed'] + analyze.confirmed['Unconfirmed'])
        self.assertEqual(analyze.counteroptions['C1'], analyze.counteroptions['C'] - analyze.counteroptions['C2'])
        #####

if __name__ == '__main__':
    unittest.main()