## IMPORTED MODULES AND PACKAGES
import math
import xlrd3 as xlrd
import xlsxwriter 
from datetime import datetime
from datetime import timedelta


# File name that contains information that is being read
datafile="Commission spreadsheet.xls"

rbook = xlrd.open_workbook(datafile)


## Rounding function to accurately round positive decimal numebers
def round_half_up(n, decimals):
    multiplier = 10 ** decimals
    return math.floor(n*multiplier + 0.5) / multiplier



## Questions asked in the beginning
date=input("Enter the Date (YYYY-MM-DD): ")
USD_to_CAD=0.85
EUR_to_CAD=0.77
EUR_to_USD=0.9
CAD_to_USD=1.21
#USD_to_CAD=float(input("Enter the USD to CAD exchange rate: "))
#EUR_to_CAD=float(input("Enter the EUR to CAD exchange rate: "))
#EUR_to_USD=float(input("Enter the EUR to USD exchange rate: "))
#CAD_to_USD=float(input("Enter the CAD to USD exchange rate: "))


number_of_sheets= int(input("Enter the number of names you want to enter: "))



## Determines the full name of the sales person
def name_determinant(name):
    if name == "John F":
        return "John Fallis"
    elif name == "Jacek":
        return "Jacek Zdienicki"
    elif name == "Ed C":
        return "Ed Cusack"
    elif name == "Carlos":
        return "Carlos Cruz"
    elif name == "Dan":
        return "Dan Wickman"
    elif name == "Scot":
        return "Scot Chin"
    elif name == "Aaron":
        return "Aaron Black"
    else:
        return name



## Determines the exchange rate
def ex_determinant(ex):
    first= ex[0:3]
    second= ex[-3:]
    if first == second:
        return [1,first,second]

    elif first == "EUR":
        if second == "USD":
            return [EUR_to_USD,first,second]
        else:
            return [EUR_to_CAD,first,second]

    elif first == "USD":
        return [USD_to_CAD,first,second]

    else:
        return [CAD_to_USD,first,second]



# File name of where the output will be
output_file= "Results.xlsx"

wbook= xlsxwriter.Workbook(output_file) 


## Main Formatting of Output Excel Sheet
align_c= wbook.add_format()
align_c.set_align('center')
align_l= wbook.add_format()
align_l.set_align('left')
align_r= wbook.add_format()
align_r.set_align('right')

bold= wbook.add_format({"bold":True,"underline":True})
merge= wbook.add_format({"bold":True,"underline":True})
merge.set_align('center')
merge.set_font_size(12)
merge.set_font_name('Times New Roman')
bold.set_font_size(12)
bold.set_font_name('Times New Roman')
currency_format = wbook.add_format({'num_format': '$#,##0.00'})

writing= wbook.add_format()
writing.set_font_size(11)
writing.set_font_name('Calibri')
bold_only= wbook.add_format({"bold":True})
bold_only.set_align('left')
bottom_border= wbook.add_format({"bottom":1})



## Functions that assist in reading the data from the Master Sheet
sheet= rbook.sheet_by_index(0)


def data_list(sheet):
    data= []
    row=0
    for row in range(sheet.nrows):
        col=0
        elem_data=[]
        for col in range(sheet.ncols):
            elem_data= elem_data + [sheet.cell_value(row,col)]
            col= col+1
        data= data + [elem_data]
        row= row + 1
    return data
        
lol_data= data_list(sheet)[7:]


def cust_rows(data):
    count=0
    for i in range(len(data)-1):
        i=i+1
        if data[i][0] == "":
            count=count+1
        else:
            return count+1
    return count+1



def serial_nums(data):
    res=[]
    for i in range(len(data)):
        if data[i][5] != "":
            if type(data[i][5]) != str:
                res= res + [round(data[i][5])]
            else:
                res= res + [data[i][5]]
    return res


def serial_output(serials):
    res=""
    for i in range(len(serials)):
        res= res + str(serials[i]) + ", "

    return res[:-2]


def deposits(data):
    col= 8
    deps=[]
    num=0
    while col < sheet.ncols:
        if data[0][col] != "":
            num= num + 1
            deps= deps + [[data[0][col],num]]
            col = col + 3
        else:
            col=col+1
    return deps


def date_recd(deposit,data):
    ind= 3*(deposit[1] - 1) + 9
    date= datetime.strptime(data[0][ind], "%Y-%m-%d")
    return date


def sales_rep(data):
    info=[]
    for i in range(len(data)):
        if data[i][1] != "":
            info=info + [data[i][1]]
            i=i+1
        else:
            i=i+1
            
    return info


def rep_calc(info):
    ind= info.index("@")
    rep= info[:ind-1]
    return rep


def commisions_lst(info):
    res=[]
    ind= info.index("@") + 2
    string= info[ind:]
    count= string.count(",")
    i=0
    while True:
        s= string[i:]
        if len(s) <= 0:
            return res
        ind= s.index("%")
        res= res + [float(s[:ind])]
        i= i + ind + 2

    return res


def break_amt(info):
    if info == "":
        return []
    elif type(info) == float:
        return [info]
    else:
        res= info.split(" ")
        for i in range(len(res)):
            res[i]= float(res[i])
            i+=1
        return res


def dep_sum(deps,num):
    res=0
    for i in range(num):
        res= res + deps[i][0]
    return res


def breaks_com_calc(com_lst,breaks,sums,amount):
    length= len(breaks)

    for i in range(length):
        checker= sums-breaks[i]
        
        if checker <= 0:
            res= [com_lst[i]*amount/100,str(com_lst[i]) + "%"]
            return [res,1]

        elif checker >= amount:
            i=i+1
            if i == length:
                res= [com_lst[i]*amount/100,str(com_lst[i]) + "%"]
                return [res,1]

        else:
            return [ [str(com_lst[i]) + "%",com_lst[i]*(amount-checker)/100,amount-checker],
                     [str(com_lst[i+1]) + "%",com_lst[i+1]*checker/100,checker] ]




## Title of the file
def title_output():
    excel_sheet.merge_range('A9:I9', 'COMMISIONS PAID', merge) 



## Functions that determine where the peice of information will go on the file
def Payments_received_txt(row,num,payment,exchange,com_str,c1,c2,com_amt):
    excel_sheet.write(row,1,"Payment #{0}".format(num),bold)
    excel_sheet.write(row,4,payment,currency_format)
    excel_sheet.write(row,5,"@",align_c)
    excel_sheet.write(row,6,com_str)
    excel_sheet.write(row,7,round_half_up(com_amt,2),currency_format)
    excel_sheet.write(row,8,c1,align_l)
    excel_sheet.write(row+1,4,"exchange")
    excel_sheet.write(row+1,6,exchange)
    ex= exchange*com_amt
    excel_sheet.write(row+1,7,round_half_up(ex,2),currency_format)
    excel_sheet.write(row+1,8,c2)
    row=row+2

def Machine_txt_Output(row,customer,serial_num,payment_num,payment,exchange,com_str,c1,c2,com_amt):
    excel_sheet.write(row,0,"Machine Sold To: ",bold)
    excel_sheet.write(row,2,customer)
    excel_sheet.write(row,6,"Serial # ",bold)
    excel_sheet.write(row,7,serial_num,align_l)
    row=row+2
    for i in range(payment_num):
        Payments_received_txt(row,payment_num,payment,exchange,com_str,c1,c2,com_amt)
        i=i+1



def info_by_customer(data,customer,date):
    deps= deposits(data)
    res=[]
    num= []
    for n in range(len(deps)):
        if (date - timedelta(days=11)) <= date_recd(deps[n],data) <= date:
            res= res + [deps[n][0]]
            num= num + [n]
            n=n+1
        else:
            n=n+1

    return [res,num]



## OUTPUTTING INFO AND FINAL FUNCTIONS
def output_info(row,name,date,lol_data,total_com,exchange,curr):
    ind = cust_rows(lol_data)
    data= lol_data[0:cust_rows(lol_data)]
    if len(data) <= 0:
        return [total_com, row, exchange, curr]

    else:
        rep_info= sales_rep(data)
        for i in range(len(rep_info)):
            if rep_calc(rep_info[i]) == name:
                customer= data[0][0]
                    
                exchange= ex_determinant(data[i+1][2])[0]
                curr_1 = ex_determinant(data[i+1][2])[1]
                curr_2 = ex_determinant(data[i+1][2])[2]
                curr = curr_2
                    
                serial_lst= serial_nums(data)
                serials= serial_output(serial_lst)
                formatted_date= datetime.strptime(date, "%Y-%m-%d")

                deps= info_by_customer(data,customer,formatted_date)[0]
                numbers= info_by_customer(data,customer,formatted_date)[1]
                    
                commission_lst= commisions_lst(rep_info[i])

                if len(commission_lst) == 1:
                    commission= commission_lst[0]
                    for p in range(len(deps)):
                        amount= deps[p]
                        num= numbers[p] + 1

                        com_amt= amount*commission/100
                        com_str=str(commission)+"%"
                            
                        total_com= total_com + (com_amt * exchange)
                        Machine_txt_Output(row,customer,serials,num,amount,exchange,
                                            com_str,curr_1,curr_2,com_amt)
                        row=row+5

                else:
                    breaks= break_amt(data[i+1][3])
                    all_deps= deposits(data)
                    for p in range(len(deps)):
                        amount= deps[p]
                        num= numbers[p] + 1
                        all_deps= deposits(data)
                        sums= dep_sum(all_deps,num)

                        com_num= breaks_com_calc(commission_lst,breaks,sums,amount)[1]

                            

                        if com_num == 1:
                            com_amt= breaks_com_calc(commission_lst,breaks,sums,amount)[0][0]
                            com_str= breaks_com_calc(commission_lst,breaks,sums,amount)[0][1]
                            total_com= total_com + (com_amt * exchange)
                            Machine_txt_Output(row,customer,serials,num,amount,exchange,
                                                com_str,curr_1,curr_2,com_amt)
                            row=row+5
                            
                        else:
                            for s in range(2):
                                com_str= breaks_com_calc(commission_lst,breaks,sums,amount)[s][0]
                                com_amt= breaks_com_calc(commission_lst,breaks,sums,amount)[s][1]
                                break_amount= breaks_com_calc(commission_lst,breaks,sums,amount)[s][2]
                                total_com= total_com + (com_amt * exchange)
                                Machine_txt_Output(row,customer,serials,num,break_amount,exchange,
                                                    com_str,curr_1,curr_2,com_amt)
                                s=s+1

                                row=row+5
                
            else:
                i=i+1

    return output_info(row,name,date,lol_data[ind:],total_com,exchange,curr) 



## Main loop that creates multiple sheets based on the number of names you want to enter
for sn in range(number_of_sheets):
    
    name=input("Enter name #" + str(sn+1) +": ")

    excel_sheet= wbook.add_worksheet(name) #creating and naming the sheet


    #Sizing the columns
    excel_sheet.set_column(0,3,8.40)
    excel_sheet.set_column(4,4,11.86)
    excel_sheet.set_column(5,5,3.43)
    excel_sheet.set_column(6,6,8.33)
    excel_sheet.set_column(7,7,12.40)
    excel_sheet.set_column(8,8,8.40)

    
    excel_sheet.insert_image('A1', 'pfm.png',{'x_scale': 0.27, 'y_scale': 0.28})

    title_output()
    
    output_info(11,name,date,lol_data,0,1,"CAD")


    #Total commision, row#, and exchange rate calculated to proceed to the next lines
    total_com= output_info(11,name,date,lol_data,0,1,"CAD")[0]
    row= output_info(11,name,date,lol_data,0,1,"CAD")[1]
    ex_rate= output_info(11,name,date,lol_data,0,1,"CAD")[2]
    curr= output_info(11,name,date,lol_data,0,1,"CAD")[3]


    #Pay Date is determined by adding 7 days
    dt= datetime.strptime(date, "%Y-%m-%d")
    pay_date= (dt + timedelta(days=7))


    #Total Commission and statement of agreement text output
    excel_sheet.write(row,0,"Total Commission Paid",bold)
    excel_sheet.write(row,3,str(pay_date)[:10],writing)
    excel_sheet.write(row,5," = ",align_c)
    excel_sheet.write(row,6,round_half_up(total_com,2), currency_format)
    excel_sheet.write(row,7,curr +" Gross", bold_only)

    row=row+2

    excel_sheet.write(row,0,"I hereby accept the above information to be correct and acknowledge that as of ______________________________",writing)

    row=row+1

    full_name= name_determinant(name)

    excel_sheet.write(row,0,"the only amount matured and owed for commissions to " +
                      full_name + " is",writing)
    excel_sheet.write(row,7,round_half_up(total_com,2), currency_format)

    row= row+3


    #Signature text output
    excel_sheet.write(row,0," ",bottom_border)
    excel_sheet.write(row,1," ",bottom_border)
    excel_sheet.write(row,2," ",bottom_border)
    excel_sheet.write(row,3," ",bottom_border)
    excel_sheet.write(row,6," ",bottom_border)
    excel_sheet.write(row,7," ",bottom_border)
    excel_sheet.write(row,8," ",bottom_border)

    row= row+1
    excel_sheet.write(row,0,full_name, writing)
    excel_sheet.write(row,6,"Date", writing)

    row= row+3

    excel_sheet.write(row,0," ",bottom_border)
    excel_sheet.write(row,1," ",bottom_border)
    excel_sheet.write(row,2," ",bottom_border)
    excel_sheet.write(row,3," ",bottom_border)
    excel_sheet.write(row,6," ",bottom_border)
    excel_sheet.write(row,7," ",bottom_border)
    excel_sheet.write(row,8," ",bottom_border)

    row= row+1
    excel_sheet.write(row,0,"Daniele Bisio", writing)
    excel_sheet.write(row,6,"Date", writing)



    sn=sn+1



wbook.close()