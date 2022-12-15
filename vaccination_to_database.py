#from openpyxl import  load_workbook
import mysql.connector

while True:
    vaccine_list = ['Sinovac', 'AstraZeneca', 'Johnson & Johnson', 'Moderna', 'Sinopharm', 'Pfizer']
    try:
        print('โปรดกรอกข้อมูลตามที่กำหนด\n')
        name = input('ชื่อ-สกุล >> ')
        for i in range(len(vaccine_list)):
            print(f'{i+1}) {vaccine_list[i]}')
        vaccine = int(input('ยี่ห้อวัคซีนที่ได้รับ(เป็นตัวเลข)>> '))
        amount = int(input('จำนวนวัคซีนที่ได้รับ >> '))
    except:
        print('โปรดใส่ข้อมูลให้ถูกต้อง\n')
    else:
        # Excel
        #workbook = load_workbook('Vaccination.xlsx')
        #sheet = workbook.active

        values = []
        #for row in sheet.iter_rows(min_row=2, values_only=True):
        #   print(row)
        #    values.append(row)

        # List fill
        values.append(name)
        values.append(vaccine_list[vaccine-1])
        values.append(amount)
        print(values)
        # Database
        db = mysql.connector.connect(
            host='localhost',
            port=3306,
            user='root',
            password='123456789',
            database='vaccination_data'
        )

        cursor = db.cursor()
        sql = '''
            INSERT INTO vaccination (name, vaccine, amount)
            VALUES (%s, %s, %s);
        '''

        cursor.execute(sql, values)
        db.commit()
        print('เพิ่มข้อมูลจำนวน ' + str(cursor.rowcount) + ' แถว')

        cursor.close()
        db.close()
        break