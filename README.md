# automatemail
Automate Email Using Python.

This can be used when you have a large list of people's name and their email and you want to send same email to all of them but at the same time you want yo address each one with their name. 

To address each person with name I have created a template ```invite.txt ``.

In the ```get_contacts``` function add filepath and name of the excel sheet.
```
wb = openpyxl.load_workbook(os.path.join('#add path to file', filename))
sheet = wb.get_sheet_by_name('#name of the sheet')

```

Do the same in ``` read_template ``` function
```
read_template
filepath = os.path.join('#add path to file', filename)

```
Add your Email username and password
```
email_username = 'example@gmail.com'   #add you email username
email_password = 'example'  # add your email password

```

I have added an axample.xlsx file and an invite.txt template to give an idea of how both of these file should look like.
