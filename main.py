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

@app.route('/')
def home():
    return "<h1>Home</h1>"

app.run(port=5090)