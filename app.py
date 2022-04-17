from docxtpl import DocxTemplate
from datetime import datetime
from os import path, remove

context = dict()

def get_values():
    global context
    name = input("Name > ")
    pin = input("PIN > ")
    email = input("Email > ")
    county = input("County > ")
    district = input("District > ")
    station = county
    road = input("Road/Street > ")
    town = input("Town > ")
    building = input("Building > ")
    area = input("Tax Area > ")
    box = input("Postal Address > ")
    code = input("Postal Code > ")
    date = input("Date Issued > ")
    context = {
        'date1': datetime.now().strftime("%d/%m/%Y"),
        'date2': date,
        'pin': pin,
        'name': name,
        'email': email,
        'county': county,
        'station': station,
        'district': district,
        'road': road,
        'town': town,
        'building': building,
        'area': area,
        'box': box,
        'code': code
    }


if __name__ == "__main__":
    get_values()
    for k, v in context.items():
        if "" == v:
            print("You Missed something")
            get_values()

    doc = DocxTemplate(r"KRA_template.docx")
    doc.render(context)
    file_name = "KRA {}.docx".format(context.get('pin'))
    doc.save(file_name)
    from docx2pdf import convert
    convert(file_name)
    if path.exists(file_name):
        remove(file_name)

    print("Done!")
