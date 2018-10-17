from xlwings.tests.restapi import TestCase
import unittest
import json

# TODO: this is WIP but the script to generate the apidocs does quite a good job already in the meantime

class TestApps(TestCase):

    def test_apps(self):
        with self.client:
            response = self.client.get('/apps')
            data = json.loads(response.data)
            pids = [app['pid'] for app in data['apps']]
            self.assertEqual(response.status_code, 200)
            self.assertTrue(self.app1.pid in pids)
            self.assertTrue(self.app2.pid in pids)

    def test_app(self):
        with self.client:
            response = self.client.get(f'/apps/{str(self.app1.pid)}/')
            data = json.loads(response.data)
            self.assertEqual(response.status_code, 200)
            self.assertEqual(self.app1.pid, data['pid'])


class TestBooks(TestCase):
    def test_app_books(self):
        with self.client:
            response = self.client.get(f'/apps/{str(self.app1.pid)}/books')
            self.assertEqual(response.status_code, 200)

    def test_books(self):
        with self.client:
            response = self.client.get(f'/books')
            self.assertEqual(response.status_code, 200)


class TestBook(TestCase):
    def test_book(self):
        with self.client:
            response = self.client.get(f'/book/{self.wb1.name}')
            self.assertEqual(response.status_code, 200)


if __name__ == '__main__':
    unittest.main()
