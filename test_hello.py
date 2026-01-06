import unittest
import hello


class TestHello(unittest.TestCase):
    def test_default(self):
        self.assertEqual(hello.greet(), "Hello, World!")

    def test_name(self):
        self.assertEqual(hello.greet("Alice"), "Hello, Alice!")


if __name__ == "__main__":
    unittest.main()
