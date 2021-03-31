import math
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

# const values
FEMALE = 'F'
MALE = 'M'
x = datetime(2020, 12, 31) #Compensation_Date
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

def get_salary(row):
    return (row[SALARY_COL])

def get_gen(row):
    if(row[SEX_COL] == 'M'):
        return 'M'
    else:
        return 'F'

def get_x(row):
    return (relativedelta((x),(row[BIRTH_COL]))).years



def get_seniority(row):
    if (row[END_WORK_COL] != '-'):
        return (relativedelta((row[END_WORK_COL]),(row[START_WORK_COL]))).years
    else:
        return (relativedelta((datetime.now()),(row[START_WORK_COL]))).years


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


def get_px(W , X,t):
    if(W==67):
        rows = men.shape[0]
        for row in range(1, rows):
            row = men.iloc[row]
            if( row[1] == X+t+1):
                return row[Px_COL]
    else:
        rows = women.shape[0]
        for row in range(1, rows):
            row = women.iloc[row]
            if (row[1] == X+t+1):
                return row[Px_COL]

def get_qx(X,t):
    age = X +t
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

def get_dis(t):
    rows = discount.shape[0]
    for row in range(1, rows):
        row = discount.iloc[row]
        if (row[0] == t+1):
            return 1+ row[1]


def getPxAndQxAndDis(W,X,t):
    px = get_px(W, X, t)
    qx = get_qx(X, t)
    discountrate = get_dis(t)
    return px,qx,discountrate


def get_section_1(row):
    last_salary = get_salary(row)
    seniority = get_seniority(row)
    section14rate,years = get_section14rate(row)
    section14percent = section14rate / 100
    if (get_gen(row) == 'M'):
        W = 67
    else:
        W= 64
    X= get_x(row)
    sigma = W - X - 2
    sum=0


    if( section14percent == 0 ):
        for t in range(0,sigma):
            px, qx, discountrate = getPxAndQxAndDis(W, X, t)
            sum += last_salary * seniority * ((pow(1.4,t+0.5) * px * qx) / discountrate)

    elif (section14percent == 1) :
        for t in range(0, sigma):
            px, qx, discountrate = getPxAndQxAndDis(W, X, t)
            sum += last_salary * (seniority- years) * ((pow(1.4,t+0.5) * px * qx) / discountrate)

    else :
        for t in range(0, sigma):
            px, qx, discountrate = getPxAndQxAndDis(W, X, t)
            sum += last_salary * (years * (1-section14percent)) * ((pow(1.4,t+0.5) * px * qx) / discountrate)

    print(sum)




    return seniority

def prob_to_live_this_year(gender, age, t):
    probability = 0
    if gender == FEMALE:
        num = women.iloc[age - 17 + t][Lx_COL]
        denum = women.iloc[age - 17][Lx_COL]
        probability = num / denum
    else:
        num = men.iloc[age - 17 + t][Lx_COL]
        denum = men.iloc[age - 17][Lx_COL]
        probability = num / denum
    return probability


def prob_to_die_in_t_years(gender, age, t):
    probability = 0
    if gender == FEMALE:
        num = women.iloc[age - 17][Lx_COL] - women.iloc[age - 17 + t][Lx_COL]
        denum = women.iloc[age - 17][Lx_COL]
        probability = num / denum
    else:
        num = men.iloc[age - 17][Lx_COL] - men.iloc[age - 17 + t][Lx_COL]
        denum = men.iloc[age - 17][Lx_COL]
        probability = num / denum
    return probability



def main():
    rows = data.shape[0]
    for row in range(1, rows):
        row = data.iloc[13]
        sum = get_section_1(row)
        # print (sum)
        break
        # res_r=has_resignation_reason(row)
        # l_14 = has_law_14(row)
        # sum_iters = clac_iters(res_r, l_14, row)
        # print(f'row:{row_num} sum iters:{sum_iters}')
    # print(prob_to_live_this_year('F', 18, 2))
    # print(prob_to_die_in_t_years('F', 18, 82))


main()













