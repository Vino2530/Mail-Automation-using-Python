# Mail-Automation-using-Python
<h2> Sending email to the members those who didn't pay on the time </h2>
  In this project, we are going to send the due reminder mail to the members. So, we need to analyze the given dataset and filter the       unpaid members and send them reminder mail.

![](Due%20excel%20img.png)

We can see the there is 3 unpaid members in the given data set. Using Pandas we are going to filer the members. 
<h2> Pandas in Python </h2>
<p> Python is a great language for doing data analysis, primarily because of the fantastic ecosystem of data-centric python packages. Pandas is one of those packages and makes importing and analyzing data much easier.</p>
<h2> Jupyter Notebook </h2>
<p>Jupyter Notebooks are a powerful way to write and iterate on your Python code for data analysis. Rather than writing and re-writing an entire program, you can write lines of code and run them one at a time. Then, if you need to make a change, you can go back and make your edit and rerun the program again, all in the same window.</p>

<h1>Steps:</h1>
 1. Importing Pandas module in Jupyter notebook.

        import openpyxl as excel

2. Loading a excel file using workbook object.
We are going to send a dues reminder mail for those who didn't paid the due amount.

        wb=excel.load_workbook('duesRecords.xlsx')

3. Opening a sheet using sheet object

        sheet=wb.get_sheet_by_name('Sheet1')

4. To verify the type of sheet

        type(sheet)

5. To get the last column using max_column parameter.
max_column replaced the get_highest_column from python's old versions 

        last_col=sheet.max_column

6. We need to find the latest month of due records sheet.
Latest month is last column of given dataset.
It returns the last column value using last_col number.

        last_month=sheet.cell(row=1, column=last_col).value

7. To find the unpaid members.
max_row pararmeter using to find the length of rows .

        unpaid_mem={}

        for r in range(2,sheet.max_row+1):
            payment=sheet.cell(row=r, column=last_col).value
            if payment != 'paid':
                name=sheet.cell(row=r,column=1).value
                email=sheet.cell(row=r,column=2).value
                unpaid_mem[name]=email

8. Importing SMTP module for email accessing

        import smtplib as sm
        conn=sm.SMTP('smtp.gmail.com',587)

9. Initiate the connection 

        conn.ehlo()

10. Encryption 

        conn.starttls()

11. Restart the connection

        conn.ehlo()

12. To login to mail id

        conn.login('dummyemailidfor.test@gmail.com','my_password')

13. Here we are going to send mail to the unpaid members which we filtered earlier

        for name, email in unpaid_mem.items():
            body= "Subject: %s dues unpaid.\nDear %s,\nRecords show that you have not paid dues for %s. Please make this payment as soon                    as possible. Thank you!'" % (last_month, name, last_month) 
            print('sending mail to %s....' %email)
            mail_status=conn.sendmail('dummyemailidfor.test@gmail.com',email,body)

        if mail_status != {}:
            print('Some problem in sending mails')
        else:
            print('all done')
