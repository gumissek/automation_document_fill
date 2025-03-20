import datetime
import os
from flask import Flask, render_template, request, flash
from flask_wtf import FlaskForm
from wtforms.fields.numeric import IntegerField, FloatField
from wtforms.fields.simple import StringField, SubmitField, EmailField
from wtforms.validators import DataRequired, Length, EqualTo, NumberRange
from flask_bootstrap import Bootstrap5
import fitz
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

#author https://github.com/gumissek/automation_doc_fill

MY_MAIL_SMTP = os.getenv('MY_MAIL_SMTP', 'smtp.gmail.com')
MY_MAIL = os.getenv('MY_MAIL', 'pythonkurskurs@gmail.com')
MY_MAIL_PASSWORD = os.getenv('MY_MAIL_PASSWORD', 'svvbtqswtoxdbchw')

app = Flask(__name__)
app.config['SECRET_KEY'] = '1234'
bootstrap = Bootstrap5(app)


class DocsForm(FlaskForm):
    email = EmailField(label='Podaj maila na ktorego wyslac plik (opcjonalne)', validators=[])

    odbiorca_dane_p1 = StringField(label='Podaj nazwe odbiorcy czesc 1.',
                                   validators=[DataRequired(), Length(min=1, max=27,
                                                                      message='Pole moze zawierac maksymalnie 27 znakow')])
    odbiorca_dane_p2 = StringField(label='Podaj nazwe odbiorcy czesc 2.', validators=[Length(min=0, max=27,
                                                                                             message='Pole moze zawierac maksymalnie 27 znakow')])
    odbiorca_nr_konta = IntegerField(label='Podaj numer rachunku odbiorcy', validators=[DataRequired(),
                                                                                        NumberRange(
                                                                                            min=10000000000000000000000000,
                                                                                            max=99999999999999999999999999,
                                                                                            message='Numer rachunku odbiorcy musi miec 26 cyfr'),
                                                                                        EqualTo(
                                                                                            'odbiorca_nr_konta_repeat',
                                                                                            message='Numery konta musza byc takie same')])
    odbiorca_nr_konta_repeat = IntegerField(label='Podaj numer rachunku odbiorcy ponownie', validators=[DataRequired(),
                                                                                                        NumberRange(
                                                                                                            min=10000000000000000000000000,
                                                                                                            max=99999999999999999999999999,
                                                                                                            message='Numer rachunku odbiorcy musi miec 26 cyfr'), ])
    kwota_liczba = FloatField(label='Podaj kwote w formacie 123.12',
                              validators=[DataRequired('Niepoprawny format danych'), NumberRange(max=999999999.99,
                                                                                                 message='Kwota z przedzialu od 0 do 999,999,999.99')])

    kwota_slownie = StringField(label='Podaj kwote slownie', validators=[DataRequired(), Length(max=27,
                                                                                                message='Pole moze zawierac maksymalnie 27 znakow')])

    nadawca_dane_p1 = StringField(label='Podaj nazwe nadawcy czesc 1.',
                                  validators=[DataRequired(), Length(min=1, max=27,
                                                                     message='Pole moze zawierac maksymalnie 27 znakow')])
    nadawca_dane_p2 = StringField(label='Podaj nazwe nadawcy czesc 2.',
                                  validators=[Length(min=0, max=27,
                                                     message='Pole moze zawierac maksymalnie 27 znakow')])
    tytul_p1 = StringField(label='Podaj tytul transakcji czesc 1.', validators=[DataRequired(), Length(min=1, max=27,
                                                                                                       message='Pole moze zawierac maksymalnie 27 znakow')])
    tytul_p2 = StringField(label='Podaj tytul transakcji czesc 2.', validators=[Length(min=0, max=27,
                                                                                       message='Pole moze zawierac maksymalnie 27 znakow')])
    submit = SubmitField(label='Wypelnij')


input_pdf = 'druk.pdf'
output_pdf = 'druk_wypelniony.pdf'


@app.route('/', methods=['POST', 'GET'])
def home():
    form = DocsForm()
    if form.validate_on_submit():
        email = request.form['email'].strip()

        odbiorca_dane_p1 = request.form['odbiorca_dane_p1'].strip().upper()
        odbiorca_dane_p2 = request.form['odbiorca_dane_p2'].strip().upper()
        odbiorca_nr_konta = request.form['odbiorca_nr_konta'].strip().replace(' ','')
        kwota_liczba = str(round(float(request.form['kwota_liczba'].strip()), 2))
        kwota_slownie = request.form['kwota_slownie'].strip().upper()
        nadawca_dane_p1 = request.form['nadawca_dane_p1'].strip().upper()
        nadawca_dane_p2 = request.form['nadawca_dane_p2'].strip().upper()
        tytul_p1 = request.form['tytul_p1'].strip().upper()
        tytul_p2 = request.form['tytul_p2'].strip().upper()

        fontsize_druk = 9
        fontsize_pokwitowanie = 7.2
        x_pokwitowanie = 155
        x = 317
        x_step = 14.2
        line_step = 24.3
        line1 = 170
        line2 = line1 + 1 * line_step

        line3 = line1 + 2 * line_step
        line4 = line1 + 3 * line_step
        line5 = line1 + 4 * line_step
        line6 = line1 + 5 * line_step
        line7 = line1 + 6 * line_step
        line8 = line1 + 7 * line_step
        line9 = line1 + 8 * line_step
        # uzupelnianie formularza


        # otwieram pdf
        doc = fitz.open(input_pdf)

        # print(f'Ilosc stron pdfa {doc.page_count}')

        # otwieram strone
        page = doc.load_page(0)
        # wpisuje tekst
        # 1 linia odbiorca dane p.1
        for index in range(len(odbiorca_dane_p1)):
            page.insert_text((x + index * x_step, line1), odbiorca_dane_p1[index], fontsize=fontsize_druk)
        # 2 linia odbiorca dane p.2
        for index in range(len(odbiorca_dane_p2)):
            page.insert_text((x + index * x_step, line2), odbiorca_dane_p2[index], fontsize=fontsize_druk)
        # 3 linia odbiorca nr konta
        for index in range(len(odbiorca_nr_konta)):
            page.insert_text((x + index * x_step, line3), odbiorca_nr_konta[index], fontsize=fontsize_druk)
        # 4 linia kwota
        for index in range(len(kwota_liczba)):
            page.insert_text((x + 15 * x_step + index * x_step, line4), kwota_liczba[index], fontsize=fontsize_druk)
        # 5 linia kwota slownie
        for index in range(len(kwota_slownie)):
            page.insert_text((x + index * x_step, line5), kwota_slownie[index], fontsize=fontsize_druk)
        # 6 linia nadawca_dane p.1
        for index in range(len(nadawca_dane_p1)):
            page.insert_text((x + index * x_step, line6), nadawca_dane_p1[index], fontsize=fontsize_druk)
        # 7 linia nadawca_dane p.2
        for index in range(len(nadawca_dane_p2)):
            page.insert_text((x + index * x_step, line7), nadawca_dane_p2[index], fontsize=fontsize_druk)
        # 8 linia tytul p.1
        for index in range(len(tytul_p1)):
            page.insert_text((x + index * x_step, line8), tytul_p1[index], fontsize=fontsize_druk)
        # 9 linia tytul p.2
        for index in range(len(tytul_p2)):
            page.insert_text((x + index * x_step, line9), tytul_p2[index], fontsize=fontsize_druk)

        # pokwitowanie

        # nr konta
        page.insert_text((x_pokwitowanie, line1), odbiorca_nr_konta, fontsize=9)
        # odbiorca p.1
        page.insert_text((x_pokwitowanie, line2), odbiorca_dane_p1, fontsize=fontsize_pokwitowanie)
        # odbiorca p.2
        page.insert_text((x_pokwitowanie, line2 + 12), odbiorca_dane_p2, fontsize=fontsize_pokwitowanie)
        # kwota
        page.insert_text((x_pokwitowanie + 20, line4 + 3), f'{kwota_liczba} PLN', fontsize=fontsize_pokwitowanie)
        page.insert_text((x_pokwitowanie, line4 + 12), kwota_slownie, fontsize=fontsize_pokwitowanie)
        # nadawca dane p.1
        page.insert_text((x_pokwitowanie, line6), nadawca_dane_p1, fontsize=fontsize_pokwitowanie)
        # nadawca dane p.2
        page.insert_text((x_pokwitowanie, line6 + 12), nadawca_dane_p2, fontsize=fontsize_pokwitowanie)
        # tytul p.1
        page.insert_text((x_pokwitowanie, line8), tytul_p1, fontsize=fontsize_pokwitowanie)
        # tytul p.2
        page.insert_text((x_pokwitowanie, line8 + 12), tytul_p2, fontsize=fontsize_pokwitowanie)
        # zapisuje plik
        doc.save(output_pdf)
        flash(f'Dokument zostal wygenerowany {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}')

        # wysylanie emailem zalacznika
        if email != '':
            subject = f'Wygenerowany dokument z dnia {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}'
            body = 'Dokument:'

            msg = MIMEMultipart()
            msg['From'] = MY_MAIL
            msg['To'] = email
            msg['Subject'] = subject

            msg.attach(MIMEText(body, 'plain'))

            with open(output_pdf, 'rb') as zalacznik:  # otwarcie pliku w trybie binarnym
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(zalacznik.read())  # odczytujemy zawartosc pliku
                encoders.encode_base64(part)
                part.add_header('Content-Disposition',
                                f'attachment; filename={output_pdf.split('/')[-1]}')  # dodaje naglowek zalcznika z nazwa pliku
                msg.attach(part)  # dolaczam plik do wiadomosci

            tresc_maila = msg.as_string()

            with smtplib.SMTP(MY_MAIL_SMTP, port=587) as connection:
                connection.starttls()
                connection.login(user=MY_MAIL, password=MY_MAIL_PASSWORD)

                connection.sendmail(from_addr=MY_MAIL, to_addrs=email,
                                    msg=tresc_maila)
            flash(f'Wiadomosc z dokumentem wyslana na email {email}.')

        else:
            os.system(f'open {output_pdf}')

    return render_template('home.html', form=form)


if __name__ == '__main__':
    app.run(debug=False, port=5001)
