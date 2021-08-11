import openpyxl as xl
import os
import csv
from scavenger import tools


DOCROOT = "/KoreaBase/bp6/scavenger/database"

def flatten(t):
    return [item for sublist in t for item in sublist]


def import_success():
    return "Hello, world!"


def client(print=False):
    """Search for client related xlsx files, and return a list containing them."""
    
    fields_list = [
        ["사업체명"], ["대표자"], ["사업자번호"], ["업태/종목"],
        ["세금계산서이메일"], ["주소"], ["전화번호"], ["담당자/연락처"]
    ]
    files = tools.get_files_under(
        os.path.join(DOCROOT, "source/CLIENT/"),
        excludes=["견적서", "급여", "Fin 합", "Payment Summary"],
        oldest=2020)
    with open(os.path.join(DOCROOT, "output", "client.csv"), 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=flatten(fields_list)+["담당자", "연락처"])
        writer.writeheader()
        for f in files:
            data = tools.extract_horizontal(f, fields_list, ["새 계약서 폼"], range_condition='"갑"')
            if "담당자/연락처" in data and '/' in data["담당자/연락처"]:
                raw = data["담당자/연락처"].split('/')
                data["담당자"] = raw[0]
                data["연락처"] = raw[1]
                del data["담당자/연락처"]
            if print:
                print(data)
            writer.writerow(data)

def contract():
    pass

def quote():
    pass

def location():
    pass

def worker():
    pass

def assignment():
    pass

def salary():
    pass

if __name__ == "__main__":
    client()