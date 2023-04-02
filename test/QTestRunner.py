import unittest
from QTestAdminLoginScreen import AdminLoginTest
from QTestAdminHomeScreen import AdminHomeTest

class QTestRunner:

    testList = [
        AdminLoginTest,
        AdminHomeTest
        ]
    testLoad = unittest.TestLoader()

    TestList = []
    for testCase in testList:
        testSuite = testLoad.loadTestsFromTestCase(testCase)
        TestList.append(testSuite)
    
    newSuite = unittest.TestSuite(TestList)
    runner = unittest.TextTestRunner()
    runner.run(newSuite)