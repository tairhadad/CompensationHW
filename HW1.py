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

dismissal_resignation = {"A": {"min_age": 18, "max_age": 29, "dismissal": 0.12, "resignation": 0.17},
                         "B": {"min_age": 30, "max_age": 39, "dismissal": 0.08, "resignation": 0.1},
                         "C": {"min_age": 40, "max_age": 49, "dismissal": 0.05, "resignation": 0.1},
                         "D": {"min_age": 50, "max_age": 59, "dismissal": 0.04, "resignation": 0.07},
                         "E": {"min_age": 60, "max_age": 67, "dismissal": 0.02, "resignation": 0.02}}

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
        return 67
    else:
        return 64

def get_x(row):
    return (relativedelta((x),(row[BIRTH_COL]))).years



def get_seniority(row):
    if (row[END_WORK_COL] != '-'):
        return (relativedelta((row[END_WORK_COL]),(row[START_WORK_COL]))).years
    else:
        return (relativedelta((datetime.now()),(row[START_WORK_COL]))).years

def get_section14percent(row):
    return (row[LAW_14_percent_COL] / 100)


def get_section14rate(row):
    if (str(row[LAW_14_COL]) == 'nan'):
        return 0,0
    else:
        untilnow = (relativedelta((datetime.now()),(row[START_WORK_COL]))).years
        until14 = (relativedelta((row[LAW_14_COL]), (row[START_WORK_COL]))).years
        since14 = untilnow-until14
        return until14,since14

def get_section_1(row):
    last_salary = get_salary(row)
    seniority = get_seniority(row)
    section14rate = get_section14rate(row)
    section14percent = get_section14percent(row)
    W = get_gen(row)
    X= get_x(row)
    sigma = W - X - 2
    #for i in range(0,sigma):

    print (sigma)
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


def prob_to_die_this_year(gender, age, t):
    return 1 - prob_to_live_this_year(gender, age, t)


def calc_dismissal(salary, seniority, discout, p, fired):
    num = math.pow(1 + SALARY_GROWTH_RATE, GROWTH_POWER) * fired["dismissal"] * p
    denum = math.pow(1 + discout, t + GROWTH_POWER)
    return salary * seniority * (num / denum)


def calc_resignation(worth, resignation):
    return worth * resignation["resignation"]


def calc_EOL(salary, seniority, discout, p, discounting, q):
    num = math.pow(1 + SALARY_GROWTH_RATE, GROWTH_POWER) * q * p
    denum = math.pow(1 + discounting, t + GROWTH_POWER)
    return salary * seniority * (num / denum)


def has_resignation_reason(row):
    return not pd.isnull(row[RESIGNATION_COL])


def has_law_14(row):
    return not pd.isnull(row[LAW_14_COL])


def clac_iters(gender):
    total = 0
    if gender == FEMALE:
        total = W_WOMEN - row[BIRTH_COL].year - 2


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













