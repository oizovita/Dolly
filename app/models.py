from app import db

class Dolly(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    login = db.Column(db.String(20), unique=True)
    email = db.Column(db.String(20), unique=True)
    password = db.Column(db.String(35))

    def __init__(self, *args, **kwargs):
        super(Dolly, self).__init__(*args, **kwargs)


