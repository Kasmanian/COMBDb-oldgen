from model import Model

model = Model()

def test_admin_login_true():
    assert model.adminLogin('admin2', 'password'.encode('utf-8')) == True

def test_admin_login_false():
    assert model.adminLogin('admin5', 'password'.encode('utf-8')) == False

def test_assert_false():
    assert False