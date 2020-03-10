from flask import Flask, render_template, request
#This module provide the functions for interacting with excel sheet using python
import xlrd
#This module provide support to generate random number for weights in range 0 to 1 for prediction process
import random

#Calculating forcated patients
#By Using Weighted average forecast (Time series forcasting problem)
#It is appropriate because the training data includes time and value against it
def expSmoCal(num, l):
    li = list()
    for i in range(num):
        w1 = random.uniform(0, 1)
        w2 = random.uniform(0, 1 - w1)
        w3 = 1 - (w1+w2)
        F = w1*(l[(len(l))-3]) + w2*(l[(len(l))-2]) + w3*(l[(len(l))-1])
        l.append(F)
        li.append(int(F))
    return li

#Extracting weeks data from spreadsheet
#This is user defined method for extracting data from an excel sheet for using it predict next week patient visits
def extWeeksData():
    workbook = xlrd.open_workbook("data.xlsx")
    worksheet = workbook.sheet_by_index(0)
    trow = worksheet.nrows
    data = list()
    for i in range(trow):
        data.append(worksheet.cell(i, 1).value)
    return data

#Extracting months data from spreadsheet
#This is user defined method for extracting data from an excel sheet for using it predict next month patient visits
def extMonthsData():
    workbook = xlrd.open_workbook("data.xlsx")
    worksheet = workbook.sheet_by_index(1)
    trow = worksheet.nrows
    data = list()
    for i in range(trow):
        data.append(worksheet.cell(i, 1).value)
    return data

#Extracting years data from spreadsheet
#This is user defined method for extracting data from an excel sheet for using it predict next year patient visits

def extYearsData():
    workbook = xlrd.open_workbook("data.xlsx")
    worksheet = workbook.sheet_by_index(2)
    trow = worksheet.nrows
    data = list()
    for i in range(trow):
        data.append(worksheet.cell(i, 1).value)
    return data

app = Flask(__name__)

#Below is the route for home page
#This method rander index.html page
@app.route('/index.html')
@app.route('/')
def index():
    return render_template("index.html")

#Below methon rander graph page
#This method also check if user enter more than one value
@app.route('/graph.html',  methods=['GET', 'POST'])
def graph():
    Weeks = (request.form['Weeks'])
    Months = (request.form['Months'])
    Years = (request.form['Years'])
    if Weeks and not Months and not Years:
        li = extWeeksData()
        li2 = expSmoCal(int(Weeks), li)
        labels = [i for i in range(int(Weeks))]
        values = li2
        rang = len(values)

        return render_template('graph.html', values=values, labels=labels, ran=rang)
    if Months and not Weeks and not Years:
        li = extMonthsData()
        li2 = expSmoCal(int(Months), li)
        labels = [i for i in range(int(Months))]
        values = li2
        rang = len(values)

        return render_template('graph.html', values=values, labels=labels, ran=rang)
    if Years and not Weeks and not Months:
        li = extYearsData()
        li2 = expSmoCal(int(Years), li)
        labels = [i for i in range(int(Years))]
        values = li2
        rang = len(values)

        return render_template('graph.html', values=values, labels=labels, ran=rang)
    return render_template('error.html')


if __name__ == '__main__':
    app.run(debug=True)