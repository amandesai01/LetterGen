from flask import Flask, request, render_template, session, redirect
from flask_wtf import FlaskForm
from wtforms import StringField, BooleanField, IntegerField, SubmitField, SelectField, TextAreaField
from Letter import make

app = Flask(__name__)
app.config['SECRET_KEY'] = 'you-will-never-guess'

class GeneratorForm(FlaskForm):
    To = StringField('TO')
    IsAssistantComm = BooleanField('Is Assistant Commissioner?')
    WardNo = IntegerField("Ward No")
    Assessee = SelectField("Select Assessee ", choices=make.get_choices_kvp())
    Subject = StringField('Subject')
    RefNo = StringField("RefNo")
    Body = TextAreaField("Body")
    Submit = SubmitField('Generate')

class AssesseeForm(FlaskForm):
    Name = StringField("Name")
    Address = StringField("Address")
    MobNo = StringField("Mobile No")
    PAN = StringField("PAN")
    Submit = SubmitField('Add')

@app.route('/generate', methods=['GET', 'POST'])
def generate():
    form = GeneratorForm()
    if form.is_submitted():
        To = form.To.data
        IsAssistantComm = form.IsAssistantComm.data
        WardNo = form.WardNo.data
        Assessee = form.Assessee.data
        Subject = form.Subject.data
        RefNo = form.RefNo.data
        Body = form.Body.data
        letter = make.Letter(To, IsAssistantComm, WardNo, Assessee, Subject, RefNo, Body)
        letter.save_to_docs('testing.docx')
        return redirect('/')
    return render_template('main.html', form=form)

@app.route('/add_assessee')
def add_assessee():    
    form = AssesseeForm()
    if form.is_submitted():
        Name = form.Name.data
        Address = form.Address.data
        MobNo = form.MobNo.data
        PAN = form.MobNo.data
        try:
            # make.add_assessee()
            print("adding")
        except Exception as e:
            print(str(e))
        return redirect('/')
    return render_template('new_assessee.html')
@app.route('/')
def home():
    return "<h1>Home</h1>"

app.run(port=5090)