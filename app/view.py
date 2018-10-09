from app import app, db
from flask import render_template
from flask import request, redirect, url_for
from models import Dolly
import models


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/login', methods=['POST', 'GET'])
def login():
    if request.method == 'POST':
        login = request.form['login']
        password = request.form['password']

        if Dolly.query.filter(Dolly.login.contains(login)).all() and Dolly.query.filter(Dolly.password.contains(password)).all():
            return render_template('user.html', name=login)
        else:
            return render_template('login.html', message='Wrong login or password')

    return render_template('login.html')


@app.route('/registration', methods=['POST', 'GET'])
def registration():
    if request.method == 'POST':
        login = request.form['login']
        email = request.form['email']
        password = request.form['password']

        try:
            dolly = Dolly(login=login, email=email, password=password)
            db.session.add(dolly)
            db.session.commit()
        except:
            return render_template('registration.html', message="Login or email do not fit")

        return redirect(url_for('login'))

    return render_template('registration.html')
