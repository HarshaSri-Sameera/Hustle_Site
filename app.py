from flask import render_template
from flask import Flask, flash, redirect, request, session, abort
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import os
import smtplib

app = Flask(__name__)
app.secret_key = os.urandom(100)

# --------------------------------------------------------------------------------------
# GSPREAD SETUP
SCOPE = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/spreadsheets",
         "https://www.googleapis.com/auth/drive.file",
         "https://www.googleapis.com/auth/drive"]
CREDS = ServiceAccountCredentials.from_json_keyfile_name('static/cred.json', SCOPE)
client = gspread.authorize(CREDS)
Sheet = client.open_by_key('1pzCWLj0Azskd8LVggQOrvsKbM-eKN1P3u6wLpZiEkwY').sheet1


# --------------------------------------------------------------------------------------
def check_validity(username, password):
    try:
        cell = Sheet.find(username)
        if password == Sheet.cell(cell.row, cell.col + 1).value:
            return True
        else:
            return False
    except gspread.exceptions.CellNotFound:
        print("Not Found!")
        return False


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/login')
def login():
    if not session.get('logged_in'):
        return render_template('login.html')
    else:
        return redirect("/Hustler_Portal")


@app.route("/logout_process")
def logout():
    session['logged_in'] = False
    return redirect("/")


@app.route('/login_process', methods=['POST'])
def login_process():
    username = request.form.get("username")
    password = request.form.get("password")
    validity = check_validity(username, password)
    if validity:
        session['logged_in'] = True
        return login()
    else:
        flash('wrong password!')
        return login()


@app.route('/Hustler_Portal')
def portal():
    return render_template('portal.html')


@app.route('/contact', methods=['POST'])
def contact():

    name = request.form.get("name")
    email = request.form.get("email")
    subject = request.form.get("subject")
    message = request.form.get("message")
    body = f'''
Name: {name} | Email: {email} 

Subject: {subject}

Message: 
{message}
'''
    print(body)
    message = 'Subject: {}\n\n{}'.format(subject, body)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    # server.login(os.environ.get('EMAIL'),os.environ.get('PASSWORD'))
    server.login('Hustle247.Notify@gmail.com','Hustlers2021')
    server.sendmail('Hustle247.Notify@gmail.com', 'hustle247.clan@gmail.com', message)
    return render_template('index.html')


if __name__ == '__main__':
    app.run()
