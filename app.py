from flask import Flask, request, render_template, send_file
import json
import pandas as pd

app = Flask(__name__)
data_store = []  # Armazena os dados do formulário
current_id = 0  # ID inicial para registros

@app.route('/')
def form():
    return render_template('form.html')  # Página do formulário

@app.route('/submit', methods=['POST'])
def submit():
    global current_id
    user_data = request.form.to_dict()
    user_data['id'] = current_id  # Adiciona ID ao registro
    current_id += 1  # Incrementa o ID
    data_store.append(user_data)  # Adiciona ao armazenamento

    # Salva os dados no arquivo JSON
    with open('data.json', 'w') as json_file:
        json.dump(data_store, json_file, indent=4)

    return render_template('confirmation.html')  # Página de confirmação

@app.route('/view-data')
def view_data():
    # Página que exibe os dados em uma tabela
    return render_template('view.html', data=data_store)

@app.route('/download')
def download():
    # Converte os dados para um DataFrame
    df = pd.DataFrame(data_store)

    # Salva o DataFrame em um arquivo Excel formatado
    excel_file = 'dados.xlsx'
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Registros')

        # Ajusta a largura das colunas no Excel
        workbook = writer.book
        worksheet = writer.sheets['Registros']
        for column_cells in worksheet.columns:
            worksheet.column_dimensions[column_cells[0].column_letter].width = 20

    # Envia o arquivo Excel como download
    return send_file(excel_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)  # Inicia o servidor Flask