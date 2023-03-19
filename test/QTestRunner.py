import unittest
from QTestAdminLoginScreen import AdminLoginTest

class QTestRunner:

    testList = [
        AdminLoginTest
        ]
    testLoad = unittest.TestLoader()

    TestList = []
    for testCase in testList:
        testSuite = testLoad.loadTestsFromTestCase(testCase)
        TestList.append(testSuite)
    
    newSuite = unittest.TestSuite(TestList)
    runner = unittest.TextTestRunner()
    runner.run(newSuite)