#IMPORTING ALL THE MODULES
import re
import sqlite3
import urllib
from urllib.request import urlopen
import xlsxwriter
import operator
from urllib.error import HTTPError
from urllib.error import URLError
from bs4 import BeautifulSoup
#ADDING THE CHARTS AND USING THE EXCEL FILE
workbook=xlsxwriter.Workbook("d:\\submit1.xlsx")
worksheet=workbook.add_worksheet()
chart=workbook.add_chart({"type":"column","subtype":"stacked"})
chart1=workbook.add_chart({"type":"column","subtype":"stacked"})
chart2=workbook.add_chart({"type":"column","subtype":"stacked"})
chart3=workbook.add_chart({"type":"column","subtype":"stacked"})
chart4=workbook.add_chart({"type":"column","subtype":"stacked"})
#TAKING THE URLS AS INPUT
urls=["https://www.flipkart.com/","https://www.apple.com/","https://www.snapdeal.com/","https://www.shopclues.com/","https://www.jabong.com/"]
number=0 #initialising a variable number for iterating url loop
for url in urls:
    #Checking for errors in reaching the url
    try:
        f = urlopen(url)
    except HTTPError as e:
        print("THERE IS AN ERROR IN URL: ",e)
    except URLError as e:
        print("THERE IS AN ERROR IN REACHING THE WEBSITE: ",e)
    else:
        soup = BeautifulSoup(f, "html.parser")
        for script in soup(["script", "style"]):
            script.extract()
    f1 = open("D:\\article.txt", "r")    #OPENING THE FILE WITH ALL THE WORDS TO IGNORE
    da = f1.read()
    w = da.split()
    text = soup.get_text()
    lines = (line.strip() for line in text.splitlines())
    a = []

    for word in lines:
        p=word.lower()
        p1= p.split()
        a.extend(p1)


    conn = sqlite3.connect("mysql.db")

    b = []
    for letter in a:
        if letter not in w:
            b.append(letter)
        else:
            continue
    join_group = str(1 * " ").join(b)
    join_group1 = re.sub("\d+\s|\s\d+\s|\s\d+$|#|\.|:|,|'", "", join_group)  #TO IGNORE ALL THE DIGITS IN THE WEBSITE CONTENT

    b1 = join_group1.split()
    D={}
    for word in b1:
        D1 = {word: b1.count(word)}
        D.update(D1)

    D2 = dict(sorted(D.items(), key=operator.itemgetter(1), reverse=True)[:20])  #GET THE TOP 20 OCCURING KEYWORDS OF THE WEBSITE

    if number==0: #if there is no database created then this will execute else the other condition will run
        conn.execute("DROP TABLE IF EXISTS NUMBER")   #CREATING THE DATABASE IN SQLITE3
        conn.execute('''CREATE TABLE NUMBER
                   (WORD TEXT NOT NULL,
                    COUNT INT NOT NULL,
                    URL TEXT);''')
        number+=1
        for k,v in D2.items():  #INPUTTING THE KEYWORDS INTO DATABASE
            conn.execute('INSERT INTO NUMBER(WORD,COUNT,URL) VALUES(?,?,?)',(k,v,url));

        conn.commit()

    else:
        for k,v in D2.items():  #INPUTTING THE KEYWORDS INTO DATABASE
            conn.execute('INSERT INTO NUMBER(WORD,COUNT,URL) VALUES(?,?,?)',(k,v,url));

        conn.commit()

cursor = conn.execute("SELECT WORD,COUNT,URL FROM NUMBER")
for i, row in enumerate(cursor): #IMPORTING THE DATA FROM DATABASE TO EXCEL
    print("WORD: ", row[0])
    print("COUNT: ", row[1])
    print("URL: ",row[2])
    worksheet.write(i, 0, row[0])
    worksheet.write(i, 1, row[1])
    worksheet.write(i, 2, row[2])
    #CREATING THE EXCEL CHART AND FORMATTING IT PROPERLY
chart.add_series({"name": "words",'categories': '=Sheet1!A1:A20', "values": "=Sheet1!$A$1:$A$20", 'column': {"color": 'blue'}})
chart.add_series({"name": "count", "values": "=Sheet1!$B$1:$B$20", 'column': {"color": 'green'}})
chart.set_x_axis({'name': 'WORD',"values": "=Sheet1!$A$1:$A$20", 'name_font': {'bold': True, 'italic': True}})
chart.set_y_axis({'name': 'COUNT', 'name_font': {'bold': True, 'italic': True}})
chart.set_title({'name': '=Sheet1!$C$1'})
chart.set_plotarea({
    'border': {'color': 'red', 'width': 2, 'dash_type': 'dash'},
    'fill': {'color': '#FFFFC2'}
})
chart.set_legend({'font': {'size': 9, 'bold': True}})
chart1.add_series({"name": "words", "values": "=Sheet1!$A$21:$A$40",'categories': '=Sheet1!$A$21:$A$40', 'column': {"color": 'blue'}})
chart1.add_series({"name": "count", "values": "=Sheet1!$B$21:$B$40", 'column': {"color": 'green'}})
chart1.set_plotarea({
    'border': {'color': 'red', 'width': 2, 'dash_type': 'dash'},
    'fill': {'color': '#FFFFC2'}
})
chart1.set_x_axis({'name': 'WORD', 'name_font': {'bold': True, 'italic': True}})
chart1.set_title({'name': '=Sheet1!$C$21'})
chart1.set_legend({'font': {'size': 9, 'bold': True}})
chart1.set_y_axis({'name': 'COUNT', 'name_font': {'bold': True, 'italic': True}})
chart2.add_series({"name": "words", "values": "=Sheet1!$A$41:$A$60",'categories': '=Sheet1!$A$41:$A$60', 'column': {"color": 'blue'}})
chart2.add_series({"name": "count", "values": "=Sheet1!$B$41:$B$60", 'column': {"color": 'green'}})
chart2.set_plotarea({
    'border': {'color': 'red', 'width': 2, 'dash_type': 'dash'},
    'fill': {'color': '#FFFFC2'}
})
chart2.set_x_axis({'name': 'WORD', 'name_font': {'bold': True, 'italic': True}})
chart2.set_title({'name': '=Sheet1!$C$41'})
chart2.set_legend({'font': {'size': 9, 'bold': True}})
chart2.set_y_axis({'name': 'COUNT', 'name_font': {'bold': True, 'italic': True}})
chart3.add_series({"name": "words", "values": "=Sheet1!$A$61:$A$80",'categories': '=Sheet1!$A$61:$A$80', 'column': {"color": 'blue'}})
chart3.add_series({"name": "count", "values": "=Sheet1!$B$61:$B$80", 'column': {"color": 'green'}})
chart3.set_plotarea({
    'border': {'color': 'red', 'width': 2, 'dash_type': 'dash'},
    'fill': {'color': '#FFFFC2'}
})
chart3.set_x_axis({'name': 'WORD', 'name_font': {'bold': True, 'italic': True}})
chart3.set_title({'name': '=Sheet1!$C$61'})
chart3.set_legend({'font': {'size': 9, 'bold': True}})
chart3.set_y_axis({'name': 'COUNT', 'name_font': {'bold': True, 'italic': True}})
chart4.add_series({"name": "words", "values": "=Sheet1!$A$81:$A$100", 'categories': '=Sheet1!$A$81:$A$100','column': {"color": 'blue'}})
chart4.add_series({"name": "count", "values": "=Sheet1!$B$81:$B$100", 'column': {"color": 'green'}})
chart4.set_plotarea({
    'border': {'color': 'red', 'width': 2, 'dash_type': 'dash'},
    'fill': {'color': '#FFFFC2'}
})
chart4.set_x_axis({'name': 'WORD', 'name_font': {'bold': True, 'italic': True}})
chart4.set_title({'name': '=Sheet1!$C$81'})
chart4.set_legend({'font': {'size': 9, 'bold': True}})
chart4.set_y_axis({'name': 'COUNT', 'name_font': {'bold': True, 'italic': True}})
worksheet.insert_chart("H7", chart)
worksheet.insert_chart("Q7", chart1)
worksheet.insert_chart("Z7", chart2)
worksheet.insert_chart("AI7", chart3)
worksheet.insert_chart("AS7", chart4)
conn.commit()
workbook.close()
print("END OF CODE")








