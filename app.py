from flask import Flask, render_template, request, redirect, flash, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_mail import Mail, Message
from datetime import datetime
from openpyxl import Workbook
import io






app = Flask(__name__)


    
app.secret_key = 'chave-secreta'

# Banco de dados SQLite
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///salas.db'
db = SQLAlchemy(app)

# Configuração de e-mail
app.config['MAIL_SERVER'] = 'smtp.gmail.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'pedrolof15@gmail.com'
app.config['MAIL_PASSWORD'] = 'Lotgh10993615'
mail = Mail(app)

# Modelo de dados
class Reserva(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    nome_sala = db.Column(db.String(50), nullable=False)
    usuario = db.Column(db.String(100), nullable=False)
    email = db.Column(db.String(100), nullable=False)
    inicio = db.Column(db.DateTime, nullable=False)
    fim = db.Column(db.DateTime, nullable=False)

# Página principal
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        nome_sala = request.form['nome_sala']
        usuario = request.form['usuario']
        email = request.form['email']
        inicio = datetime.strptime(request.form['inicio'], '%Y-%m-%dT%H:%M')
        fim = datetime.strptime(request.form['fim'], '%Y-%m-%dT%H:%M')

        if fim <= inicio:
            flash("O horário de fim deve ser depois do início!", "error")
            return render_template('index.html')

        conflitos = Reserva.query.filter(
            Reserva.nome_sala == nome_sala,
            Reserva.inicio < fim,
            Reserva.fim > inicio
        ).all()

        if conflitos:
            flash(f"A sala '{nome_sala}' já está reservada neste horário.", "error")
        else:
            nova_reserva = Reserva(
                nome_sala=nome_sala,
                usuario=usuario,
                email=email,
                inicio=inicio,
                fim=fim
            )
            db.session.add(nova_reserva)
            db.session.commit()

            # E-mail
            msg = Message('Reserva Confirmada',
                          sender='pedrolof15@gmail.com',
                          recipients=[email])
            msg.body = (f"Olá {usuario},\n\nSua reserva da sala '{nome_sala}' foi confirmada de "
                        f"{inicio.strftime('%d/%m/%Y %H:%M')} até {fim.strftime('%d/%m/%Y %H:%M')}.\n\n"
                        "Obrigado!")
            mail.send(msg)

            flash(f"Reserva realizada com sucesso! Confirmação enviada para {email}.", "success")

    return render_template('index.html')

# Nova rota para exportar as reservas em Excel
@app.route('/exportar_excel')
def exportar_excel():
    reservas = Reserva.query.all()

    wb = Workbook()
    ws = wb.active
    ws.title = "Reservas"

    # Cabeçalho
    ws.append(["ID", "Sala", "Usuário", "Email", "Início", "Fim"])

    for r in reservas:
        ws.append([
            r.id,
            r.nome_sala,
            r.usuario,
            r.email,
            r.inicio.strftime('%d/%m/%Y %H:%M'),
            r.fim.strftime('%d/%m/%Y %H:%M')
        ])

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return send_file(
        file_stream,
        as_attachment=True,
        download_name='reservas.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# Execução
if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True)
