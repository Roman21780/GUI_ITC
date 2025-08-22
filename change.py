with open('requirements.txt', 'r', encoding='utf-16') as f:
    content = f.read()

modified_content = content.replace('=', '==')

with open('requirements.txt', 'w', encoding='utf-8') as file:
    # Записываем измененное содержимое обратно в файл
    file.write(modified_content)

print("Замена завершена. Символ '=' заменен на '=='.")