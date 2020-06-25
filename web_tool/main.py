from flask import Flask,render_template,flash,redirect,url_for,session,logging,request
from wtforms import Form,StringField,PasswordField,TextAreaField,validators
from flask_sqlalchemy import SQLAlchemy
import urllib
from functools import wraps
import pyodbc
import dimod
from flask_table import Table,Col,LinkCol

app = Flask(__name__)

app.secret_key = "web_tool"


class UpdateForm(Form):
    name = StringField("İsim")
    last_name = StringField("Soyisim")
    username = StringField("Kullanıcı Adı")
    password = StringField("Şifre")

class RegisterForm(Form):
    name = StringField("İsim")
    last_name = StringField("Soyisim")
    username = StringField("Kullanıcı Adı")
    password = PasswordField("Şifre")
    status = StringField("Bu Kullanıcı Admin Mi?")

class OrderForm(Form):
    sp1 = StringField("Öğe 1")
    sp2 = StringField("Öğe 2")
    sp3 = StringField("Öğe 3")
    sp4 = StringField("Öğe 4")
    hm1 = StringField("Hammadde 1")
    hm2 = StringField("Hammadde 2")
    hm3 = StringField("Hammadde 3")
    hm4 = StringField("Hammadde 4")

class LoginForm(Form):
    username = StringField("Username")
    password = PasswordField("Password")

conn = pyodbc.connect("DRIVER={SQL Server};SERVER=VEEAMBACKUP\VEEAMSQL2016;PORT=1433;DATABASE=web_server;UID=sa;PWD=yasinc;")
cursor = conn.cursor()

@app.route("/")
def welcomepage():
    if "logged_in" in session:
        return render_template("index.html")
    else:
        return render_template("welcomepage.html")

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("welcomepage"))

@app.route("/login", methods = ["GET", "POST"])
def login():
    form = LoginForm(request.form)
    if request.method == "POST":
        u_name = form.username.data
        password_entered = form.password.data

        cursor = conn.cursor()

        query = "SELECT password FROM users WHERE username=?"
        cursor.execute(query,[u_name])
        password = cursor.fetchall()
        for row in password:
            if row[0] == password_entered:
                cursor = conn.cursor()
                commandstring = "SELECT name,last_name,is_admin FROM users WHERE username=?"
                cursor.execute(commandstring,[u_name])
                data = cursor.fetchall()
                for elements in data:
                    session["name"] = elements[0]
                    session["last_name"] = elements[1]
                    session["logged_in"] = True
                    session["is_admin"] = elements[2]
                    return redirect(url_for("index"))
    else:
        return render_template("login.html", form = form)

def login_required(f):
    @wraps(f)
    def decorated_function(*args,**kwargs):
        if "logged_in" in session:
            return f(*args, **kwargs)
        else:
            flash("Bu Araçlara Ulaşabilmek İçin Giriş Yapmalısınız!","danger")
            return redirect(url_for('login'))
    return decorated_function

@app.route("/index")
@login_required
def index():
    return render_template("index.html")

@app.route("/calculations")
@login_required
def calculations():
    return render_template("calculations.html")

@app.route("/deleteuser/<string:username>")
@login_required
def user(username):
    if session["is_admin"] == "yes":
        cursor = conn.cursor()
        query = "DELETE FROM users WHERE username=?"
        cursor.execute(query,[username])
        cursor.commit()
        return redirect(url_for("useradministration"))
    else:
        flash("Bu sayfaya Sadece yetkili kullanıcılar girebilir. Lütfen denemeyiniz!","danger")
        return redirect(url_for("index"))

@app.route("/useradministration")
@login_required
def useradministration():
    if session["is_admin"] == "yes":
        cursor.execute("select id,name,last_name,username,password from users")
        data = cursor.fetchall() 
        return render_template("useradministration.html",value=data)      
    else:
        flash("Bu sayfaya Sadece yetkili kullanıcılar girebilir. Lütfen denemeyiniz!","danger")
        return redirect(url_for("index"))
        

@app.route("/adduser", methods = ["GET", "POST"])
@login_required
def adduser():
    form = RegisterForm(request.form)
    if session["is_admin"] == "yes":
        if request.method == "GET":
            return render_template("adduser.html", form = form)
        else:
            cursor = conn.cursor()
            query = "INSERT INTO users (name,last_name,username,password,is_admin) VALUES (?,?,?,?,?)"
            name = form.name.data
            last_name = form.last_name.data
            username = form.username.data
            password = form.password.data
            is_admin = form.status.data
            cursor.execute(query,[name,last_name,username,password,is_admin])
            cursor.commit()
            return redirect(url_for("useradministration"))
    else:
        return redirect(url_for("index"))

@app.route("/updateuser/<string:id>", methods = ["GET", "POST"])
@login_required
def updateuser(id):
    if session["is_admin"] == "yes":
        form = UpdateForm(request.form)
        if request.method == "GET":
            cursor = conn.cursor()
            query = "SELECT * FROM users WHERE id =?"
            cursor.execute(query,[id])
            info = cursor.fetchone()
            form.name.data = info[1]
            form.last_name.data = info[2]
            form.username.data = info[3]
            form.password.data = info[4]
            return render_template("updateuser.html", form = form)
        else:
            name = form.name.data
            last_name = form.last_name.data
            username = form.username.data
            password = form.password.data
            cursor = conn.cursor()
            query = "UPDATE users SET name=?,last_name=?,username=?,password=? WHERE id=?"
            cursor.execute(query,[name,last_name,username,password,id])
            return redirect(url_for("useradministration"))
    else:
        return redirect(url_for("index"))
@app.route("/otomaticorder")
def order():
    form = OrderForm(request.form)
    return render_template("otomaticorder.html", form = form)
    


if __name__ == "__main__":
    app.run(port=7000, host=("192.168.10.44"),debug=True)