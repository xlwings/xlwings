from xlwings.tests.restapi import TestCase
import unittest
import json


class TestRestApi(TestCase):
    def get_book_urls(self, endpoint):
        return [
            f'/apps/{self.app1.pid}/books/{self.wb1.name}' + endpoint,
            f'/apps/{self.app1.pid}/books/0' + endpoint,
            f'/books/{self.wb1.name}' + endpoint,
            f'/books/0' + endpoint,
            f'/book/{self.wb1.name}' + endpoint,
            f'/book/{self.wb1.fullname}' + endpoint
        ]

    def test_get_apps(self):
        with self.client:
            response = self.client.get('/apps')
            data = json.loads(response.data)
            pids = [app['pid'] for app in data['apps']]
            self.assertEqual(response.status_code, 200)
            self.assertTrue(self.app1.pid in pids)
            self.assertTrue(self.app2.pid in pids)

    def test_get_app(self):
        with self.client:
            response = self.client.get(f'/apps/{str(self.app1.pid)}')
            data = json.loads(response.data)
            self.assertEqual(response.status_code, 200)
            self.assertEqual(self.app1.pid, data['pid'])

    def test_get_books(self):
        with self.client:
            for url in [f'/apps/{str(self.app1.pid)}/books', '/books']:
                response = self.client.get(url)
                data = json.loads(response.data)
                book_names = [book['name'] for book in data['books']]
                self.assertEqual(response.status_code, 200)
                self.assertTrue('Book2' in book_names)
                self.assertTrue('Book1.xlsx' in book_names)

    def test_get_book(self):
        for url in self.get_book_urls(''):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], self.wb1.name)

    def test_get_sheets(self):
        with self.client:
            for url in self.get_book_urls('/sheets'):
                response = self.client.get(url)
                data = json.loads(response.data)
                sheet_names = [book['name'] for book in data['sheets']]
                self.assertEqual(response.status_code, 200)
                self.assertTrue('Sheet1' in sheet_names)

    def test_get_sheet(self):
        with self.client:
            for url in self.get_book_urls('/sheets/sheet1'):
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual('Sheet1', data['name'])

    def test_get_range(self):
        with self.client:
            for url in self.get_book_urls('/sheets/sheet1/range'):
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertTrue('address' in data)

    def test_get_range_address(self):
        with self.client:
            for url in self.get_book_urls('/sheets/sheet1/range/A1:B2'):
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertTrue('address' in data)
                self.assertTrue('value' in data)

    def test_get_book_names(self):
        for url in self.get_book_urls('/names'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                names = [name['name'] for name in data['names']]
                self.assertEqual(response.status_code, 200)
                self.assertTrue('myname2' in names)
                self.assertTrue('Sheet1!myname1' in names)

    def test_get_book_name(self):
        name = 'myname2'
        for url in self.get_book_urls(f'/names/{name}'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], name)

    def test_get_book_name_range(self):
        name = 'myname2'
        for url in self.get_book_urls(f'/names/{name}/range'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], name)
                self.assertEqual(data['address'], '$A$1')

    def test_get_sheet_names(self):
        for url in self.get_book_urls('/sheets/sheet1/names'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                names = [name['name'] for name in data['names']]
                self.assertEqual(response.status_code, 200)
                self.assertTrue('Sheet1!myname1' in names)

    def test_get_sheet_name(self):
        name = 'Sheet1!myname1'
        for url in self.get_book_urls(f'/sheets/sheet1/names/{name}'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], name)

    def test_get_sheet_name_range(self):
        name = 'Sheet1!myname1'
        for url in self.get_book_urls(f'/sheets/sheet1/names/{name}/range'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], name)
                self.assertEqual(data['address'], '$B$2:$C$3')

    def test_get_charts(self):
        for url in self.get_book_urls('/sheets/sheet1/charts'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                names = [chart['name'] for chart in data['charts']]
                self.assertEqual(response.status_code, 200)
                self.assertTrue('Chart 1' in names)

    def test_get_chart(self):
        name = 'Chart 1'
        for url in self.get_book_urls(f'/sheets/sheet1/charts/{name}'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], name)

    def test_get_shapes(self):
        for url in self.get_book_urls('/sheets/sheet1/shapes'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                names = [shape['name'] for shape in data['shapes']]
                self.assertEqual(response.status_code, 200)
                self.assertTrue('Chart 1' in names)

    def test_get_shape(self):
        name = 'Chart 1'
        for url in self.get_book_urls(f'/sheets/sheet1/shapes/{name}'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], name)

    def test_get_pictures(self):
        for url in self.get_book_urls('/sheets/sheet1/pictures'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                names = [shape['name'] for shape in data['pictures']]
                self.assertEqual(response.status_code, 200)
                self.assertTrue('Picture 1' in names)

    def test_get_picture(self):
        name = 'Picture 1'
        for url in self.get_book_urls(f'/sheets/sheet1/pictures/{name}'):
            with self.client:
                response = self.client.get(url)
                data = json.loads(response.data)
                self.assertEqual(response.status_code, 200)
                self.assertEqual(data['name'], name)


if __name__ == '__main__':
    unittest.main()
