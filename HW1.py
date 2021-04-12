import math
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import csv

# const values
dateDay = datetime(2020, 12, 31) #Compensation_Date
SALARY_GROWTH_RATE = 0.04
GROWTH_POWER = 0.5
W_WOMEN = 64 #RETIRE_AGE_WOMEN
W_MEN = 67 #RETIRE_AGE_MAN
# data
SEX_COL = 3
BIRTH_COL = 4 #DOB_COL
START_WORK_COL = 5
SALARY_COL = 6
LAW_14_COL = 7
LAW_14_percent_COL = 8
END_WORK_COL = 11
RESIGNATION_COL = 14
ASSETS_COL = 9

# morality
Lx_COL = 2
Dx_COL = 3
Px_COL = 4
Qx_COL = 5

dismissal_resignation = {"A": {"min_age": 18, "max_age": 29, "dismissal": 0.15, "resignation": 0.20},
                         "B": {"min_age": 30, "max_age": 39, "dismissal": 0.10, "resignation": 0.13},
                         "C": {"min_age": 40, "max_age": 49, "dismissal": 0.04, "resignation": 0.10},
                         "D": {"min_age": 50, "max_age": 59, "dismissal": 0.05, "resignation": 0.07},
                         "E": {"min_age": 60, "max_age": 67, "dismissal": 0.03, "resignation": 0.03}}

t = 0

# global values
source = pd.ExcelFile('data6.xlsx')
mortality = pd.ExcelFile('Mortality_table.xlsx')
data = pd.read_excel(source, 'data')
discount = pd.read_excel(source, 'הנחות')
men = pd.read_excel(mortality, 'גברים')
women = pd.read_excel(mortality, 'נשים')

Qx = 'q(x)'

def get_q(age,g):
    if (g == 'M'):
        rows = men.shape[0]
        for row in range(1, rows):
            row = men.iloc[row]
            if (row[1] == age):
                return row[Qx_COL]
    else:
        rows = women.shape[0]
        for row in range(1, rows):
            row = women.iloc[row]
            if (row[1] == age):
                return row[Qx_COL]


def calc_p(g):
    age =18
    p_rate = []
    sum, prev_rate = dismissal_resignation['A']['dismissal']+dismissal_resignation['A']['resignation'], 1
    p_rate.append(1)
    for i in range(1,12): #18-29
        p_rate.append((1-sum)*prev_rate)
        prev_rate = p_rate[i]+get_q(age+i,g)
    sum = dismissal_resignation['B']['dismissal']+dismissal_resignation['B']['resignation']
    for i in range(12,22): #30-39
        p_rate.append((1-sum)*prev_rate)
        prev_rate = p_rate[i]+get_q(age+i,g)
    sum = dismissal_resignation['C']['dismissal']+dismissal_resignation['C']['resignation']
    for i in range(22,32): #40-49
        p_rate.append((1-sum)*prev_rate)
        prev_rate = p_rate[i]+get_q(age+i,g)
    sum = dismissal_resignation['D']['dismissal'] + \
        dismissal_resignation['D']['resignation']
    for i in range(32,42): # 50-59
        p_rate.append((1-sum)*prev_rate)
        prev_rate = p_rate[i]+get_q(age+i,g)
    sum = dismissal_resignation['E']['dismissal'] + \
        dismissal_resignation['E']['resignation']
    for i in range(42, 51): # 60-67
        p_rate.append((1-sum)*prev_rate)
        prev_rate = p_rate[i]+get_q(age+i,g)
    return p_rate



def get_salary(row):
    return (row[SALARY_COL])

def get_gen(row):
    return 'M' if row[SEX_COL] else 'F'


def get_x(row):
    return (relativedelta((dateDay),(row[BIRTH_COL]))).years

# def has_law_14(row):
#     return  not pd.isnull(row[LAW_14_COL])


# def salary_seniority(salary, seniority, row, section14rate):
#     sum = salary*seniority(row)
#     if has_law_14(row):
#         l_14_prec = (1 - (section14rate/100))
#         sum = sum * l_14_prec
#     return sum

def get_seniority(row):
    if (row[END_WORK_COL] != '-' and str(row[END_WORK_COL]) != 'nan'):
        return (relativedelta((row[END_WORK_COL]),(row[START_WORK_COL]))).years
    else:
        return (relativedelta((dateDay),(row[START_WORK_COL]))).years

def get_qx1(age):
    if(age <= 29 and age >= 18):
        return dismissal_resignation["A"]["dismissal"]
    elif (age <= 39 and age >= 30):
        return dismissal_resignation["B"]["dismissal"]
    elif (age <= 49 and age >= 40):
        return dismissal_resignation["C"]["dismissal"]
    elif (age <= 59 and age >= 50):
        return dismissal_resignation["D"]["dismissal"]
    else:
        return dismissal_resignation["E"]["dismissal"]

def get_qx2(age):
    if(age <= 29 and age >= 18):
        return dismissal_resignation["A"]["resignation"]
    elif (age <= 39 and age >= 30):
        return dismissal_resignation["B"]["resignation"]
    elif (age <= 49 and age >= 40):
        return dismissal_resignation["C"]["resignation"]
    elif (age <= 59 and age >= 50):
        return dismissal_resignation["D"]["resignation"]
    else:
        return dismissal_resignation["E"]["resignation"]

def get_qx3(W,age):
    if (W == 67):
        rows = men.shape[0]
        for row in range(1, rows):
            row = men.iloc[row]
            if (row[1] == age):
                return row[Qx_COL]
    else:
        rows = women.shape[0]
        for row in range(1, rows):
            row = women.iloc[row]
            if (row[1] == age):
                return row[Qx_COL]

def get_dis(t):
    rows = discount.shape[0]
    for row in range(1, rows):
        row = discount.iloc[row]
        if (row[0] == t+1):
            return 1+ row[1]


def get_assetsValue (row):
    return (row[ASSETS_COL])

def get_section14rate(row):
    if (str(row[LAW_14_COL]) == 'nan'):
        return 0,0
    untilnow = get_seniority(row)
    until14 = (relativedelta((row[LAW_14_COL]), (row[START_WORK_COL]))).years
    since14 = untilnow - until14
    if (row[LAW_14_percent_COL] == 100):
        return 100,since14
    else:
        return row[LAW_14_percent_COL], since14


def get_section_1(row,p_rate):
    last_salary = get_salary(row)
    seniority = get_seniority(row)
    section14rate,years = get_section14rate(row)
    section14percent = section14rate / 100
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W = 64
    X = get_x(row)

    sigma = W - X - 2
    sum = 0

    if(section14percent == 0 ):
        for t in range(0,sigma):

            px = p_rate[t]
            qx1 = get_qx1(X+t)
            discountrate = get_dis(t)
            sum += last_salary * seniority * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx1) / pow(discountrate,t+0.5))

    elif (section14percent == 1):
        for t in range(0, sigma):
            px = p_rate[t]
            qx1 = get_qx1(X+t)
            discountrate = get_dis(t)
            sum += last_salary * (seniority- years) * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx1) / pow(discountrate,t+0.5))

    else:
        for t in range(0, sigma):
            px = p_rate[t]
            qx1 = get_qx1(X+t)
            discountrate = get_dis(t)
            sum += last_salary * years * (1-section14percent) * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx1) / pow(discountrate,t+0.5))
        for t in range(0, sigma):
            px = p_rate[t]
            qx1 = get_qx1(X+t)
            discountrate = get_dis(t)
            sum += last_salary * (seniority-years) * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx1) / pow(discountrate,t+0.5))

    return sum

def get_section_2(row,p_rate):
    last_salary = get_salary(row)
    seniority = get_seniority(row)
    section14rate,years = get_section14rate(row)
    section14percent = section14rate / 100
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W = 64
    X = get_x(row)

    sigma = W - X - 2
    sum = 0

    if(section14percent == 0 ):
        for t in range(0,sigma):
            px = p_rate[t]
            qx3 = get_qx3(W,X+t)
            discountrate = get_dis(t)
            sum += last_salary * seniority * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx3) / pow(discountrate,t+0.5))

    elif (section14percent == 1):
        for t in range(0, sigma):
            px = p_rate[t]
            qx3 = get_qx3(W,X+t)
            discountrate = get_dis(t)
            sum += last_salary * (seniority- years) * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx3) / pow(discountrate,t+0.5))

    else:
        for t in range(0, sigma):
            px = p_rate[t]
            qx3 = get_qx3(W, X+t)
            discountrate = get_dis(t)
            sum += last_salary * (years * (1-section14percent)) * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx3) / pow(discountrate,t+0.5))
        for t in range(0, sigma):
            px = p_rate[t]
            qx3 = get_qx3(W, X+t)
            discountrate = get_dis(t)
            sum += last_salary * (seniority-years) * ((pow(1+SALARY_GROWTH_RATE,t+0.5) * px * qx3) / pow(discountrate,t+0.5))
    return sum

def get_section_3(row,p_rate):
    assetsValue = get_assetsValue(row)
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W = 64
    X = get_x(row)

    sigma = W - X - 2
    sum = 0

    for t in range(0, sigma):
        px = p_rate[t]
        qx2 = get_qx2(X+t+1)
        sum += assetsValue * px * qx2
    return sum

def get_section_4(row,p_rate):
    last_salary = get_salary(row)
    seniority = get_seniority(row)
    section14rate,years = get_section14rate(row)
    section14percent = section14rate / 100
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W = 64
    X = get_x(row)

    sum = 0
    if(section14percent == 0 ):
        px = p_rate[W-X-1]
        qx1 = get_qx1(W-1)
        discountrate = get_dis(W-X)
        sum += last_salary * seniority * ((pow(1+SALARY_GROWTH_RATE, (W-X+0.5)) * px * qx1) / pow(discountrate, (W-X+0.5)))

    elif (section14percent == 1):
        px = p_rate[W-X-1]
        qx1 = get_qx1(W-1)
        discountrate = get_dis(W-X)
        sum += last_salary * (seniority - years) * ((pow(1+SALARY_GROWTH_RATE, (W-X+0.5)) * px * qx1) / pow(discountrate, (W-X+0.5)))

    else:
        px = p_rate[W-X-1]
        qx1 = get_qx1(W-1)
        discountrate = get_dis(W-X)
        sum += last_salary * (years * (1-section14percent)) * ((pow(1+SALARY_GROWTH_RATE, (W-X+0.5)) * px * qx1) / pow(discountrate, (W-X + 0.5)))

        sum += last_salary * (seniority-years) * ((pow(1+SALARY_GROWTH_RATE, (W-X+0.5)) * px * qx1) / pow(discountrate, (W-X+0.5)))

    return sum

def get_section_5(row,p_rate):
    last_salary = get_salary(row)
    seniority = get_seniority(row)
    section14rate,years = get_section14rate(row)
    section14percent = section14rate / 100
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W = 64
    X = get_x(row)

    sum = 0
    if(section14percent == 0 ):
        px = p_rate[W-X-1]
        qx3 = get_qx3(W, W- 1)
        discountrate = get_dis(W-X)
        sum += last_salary * seniority * ((pow(1+SALARY_GROWTH_RATE, (W-X-1+0.5)) * px * qx3) / pow(discountrate, (W-X-1+0.5)))

    elif (section14percent == 1):
        px = p_rate[W-X-1]
        qx3 = get_qx3(W, W- 1)
        discountrate = get_dis(W-X)

        sum += last_salary * (seniority - years) * ((pow(1+SALARY_GROWTH_RATE, (W-X-1+0.5)) * px * qx3) / pow(discountrate, (W-X-1+0.5)))

    else:
        px = p_rate[W-X-1]
        qx3 = get_qx3(W, W- 1)
        discountrate = get_dis(W-X)
        sum += last_salary * (years * (1-section14percent)) * ((pow(1+SALARY_GROWTH_RATE, (W-X-1+0.5)) * px * qx3) / pow(discountrate, (W-X -1+ 0.5)))

        sum += last_salary * (seniority-years) * ((pow(1+SALARY_GROWTH_RATE, (W-X-1+0.5)) * px * qx3) / pow(discountrate, (W-X-1+0.5)))

    return sum

def get_section_6(row,p_rate):
    assetsValue = get_assetsValue(row)
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W = 64
    X = get_x(row)

    px = p_rate[W-X-1]
    qx2 = get_qx2(W-1)
    sum = assetsValue * px * qx2
    return sum

def get_section_7(row,p_rate):
    last_salary = get_salary(row)
    seniority = get_seniority(row)
    section14rate,years = get_section14rate(row)
    section14percent = section14rate / 100
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W = 64
    X = get_x(row)

    sum = 0
    if(section14percent == 0 ):
        px = p_rate[W-X-1]
        qx1 = get_qx1(W-1)
        qx2 = get_qx2(W-1)
        qx3 = get_qx3(W,W-1)
        discountrate = get_dis(W-X)
        sum += last_salary * seniority * ((pow(1+SALARY_GROWTH_RATE, (W-X)) * px * (1-qx1 - qx2 - qx3)) / pow(discountrate, (W-X)))

    elif (section14percent == 1):
        px = p_rate[W-X-1]
        qx1 = get_qx1(W-1)
        qx2 = get_qx2(W - 1)
        qx3 = get_qx3(W, W - 1)
        discountrate = get_dis(W-X)
        sum += last_salary * (seniority - years) * ((pow(1+SALARY_GROWTH_RATE, (W-X)) * px * (1-qx1 - qx2 - qx3)) / pow(discountrate, (W-X)))

    else:
        px = p_rate[W-X-1]
        qx1 = get_qx1(W-1)
        qx2 = get_qx2(W - 1)
        qx3 = get_qx3(W, W - 1)
        discountrate = get_dis(W-X)
        sum += last_salary * (years * (1-section14percent)) * ((pow(1+SALARY_GROWTH_RATE, (W-X)) * px * (1-qx1 - qx2 - qx3)) / pow(discountrate, (W-X )))

        sum += last_salary * (seniority-years) * ((pow(1+SALARY_GROWTH_RATE, (W-X)) * px * (1-qx1 - qx2 - qx3)) / pow(discountrate, (W-X)))

    return sum


def checkRetirment(row):
    if(get_gen(row) == 'M'):
        if(relativedelta((dateDay), (row[BIRTH_COL])).years > W_MEN):
            return True
    elif (get_gen(row) == 'F'):
        if(relativedelta((dateDay), (row[BIRTH_COL])).years > W_WOMEN):
            return True
    return False


def main():
    p_rate_men = calc_p('M')
    p_rate_women = calc_p('F')
    rows = data.shape[0]
    for row in range(1, rows):
        row = data.iloc[row]
        if (checkRetirment(row)):
            sum = get_seniority(row) * get_salary(row)
        else:
            p_rate = p_rate_men if get_gen(row) == 'M' else p_rate_women

            sum = get_section_1(row,p_rate) + get_section_2(row,p_rate) + get_section_3(row,p_rate) + get_section_4(row,p_rate) + get_section_5(row,p_rate)  + get_section_6(row,p_rate) + get_section_7(row,p_rate)
        with open('results.csv', 'a', newline='') as csvfile:
            fieldnames = ['first_name', 'last_name', 'results']
            writer = csv.writer(csvfile)
            writer.writerow([row[1], row[2], sum*2])
            csvfile.close()
    print("DONE")


main()
