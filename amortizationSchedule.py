from openpyxl import Workbook
wb = Workbook()

ws1 = wb.active
ws2 = wb.active
ws3 = wb.active
ws4 = wb.active
spaceHolder = wb.active

ws1.append(["Loan Amount", "Term", "Rate"])






def calculator():
    loanAmount = int(input("Please enter loan amount? "))
    print(loanAmount)
    term= (int(input("What is the term, in years, of this loan? "))*12)
    print(term)
    apr = (int(input("What is the interest rate of this loan? "))/100)
    print(apr)
    ws1.append([loanAmount, term , apr])
    spaceHolder.append([""])

    current = loanAmount
    print(current)
    monthlyRate = apr / 12
    print(monthlyRate)
    payment = (monthlyRate * loanAmount) / (1-(1+monthlyRate)**(-term))
    print(f"You will be making a monthly payment of {payment:,.2f}")
    ws2.append(["Month", "Interest", "Principle", "Balance"])
    number = 0
    for i in range(term):
        interestPart = (current * apr)/12
        principalPart = payment - interestPart
        print(f"Month {i+1}: Interest: ${interestPart:,.2f}, Principal: ${principalPart:,.2f}, Balance: ${current:,.2f}")
        interest2= round(interestPart, 2)
        principal2= round(principalPart, 2)
        balance2= round(current, 2)
        number+=1
        ws2.append([number,interest2 , principal2, balance2])
        current -= principalPart
        
    monthlyPay = interest2 + principal2
    print(f"{monthlyPay:,.2f}")
    
    spaceHolder.append([""])
    ws4.append([f'Monthly Pay = ${monthlyPay:,.2f}'])
   
    
    


call1 = calculator()

wb.save("AmortizationSchedule.xlsx")