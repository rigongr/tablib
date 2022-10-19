import src.tablib as tablib

headers = ['First Name', 'Last Name']
data = [('John', 'Doe'), ('Mike', 'Wazowski'), ('Bada', 'Bing')]

dataset = tablib.Dataset(*data, headers=headers, title='Testing Sheet - #2')


with open('test-2.xlsx', 'wb') as f:
    f.write(dataset.export('xlsx'))