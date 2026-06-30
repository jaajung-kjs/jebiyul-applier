import os

def test_tomok_exists(tomok_path):
    assert os.path.exists(tomok_path)

def test_geonchuk_exists(geonchuk_path):
    assert os.path.exists(geonchuk_path)
