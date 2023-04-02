import unittest
from QTestAdminLoginScreen import AdminLoginTest
from QTestAdminHomeScreen import AdminHomeTest
from QTestCultureOrderFormNav import CultureOrderFormNavTest
from QTestDUWLNav import DUWLNavTest

class QTestRunner:

    testList = [
        AdminLoginTest,
        AdminHomeTest,
        CultureOrderFormNavTest,
        DUWLNavTest
        ]
    testLoad = unittest.TestLoader()

    TestList = []
    for testCase in testList:
        testSuite = testLoad.loadTestsFromTestCase(testCase)
        TestList.append(testSuite)
    
    newSuite = unittest.TestSuite(TestList)
    runner = unittest.TextTestRunner()
    runner.run(newSuite)