from flask import Flask, request, jsonify, send_file
import pandas as pd
import re
import dns.resolver
import os

app = Flask(__name__)

# Fonction pour nettoyer un email
def clean_email(email):
    return email.strip().lower()

# Fonction pour valider la syntaxe d'un email
def is_valid_syntax(email):
    regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(regex, email) is not None

# Fonction pour vérifier si le domaine existe
def is_domain_valid(email):
    try:
        domain = email.split('@')[1]
        dns.resolver.resolve(domain, 'MX')
        return True
    except Exception:
        return False

# Fonction pour détecter les emails suspects (ex: no-reply)
def is_suspect_email(email):
    suspect_keywords = ['noreply', 'no-reply', 'test', 'example', 'fake', 'admin']
    return any(keyword in email for keyword in suspect_keywords)

@app.route('/clean_emails', methods=['POST'])
def clean_emails():
    # Récupérer le fichier Excel envoyé dans la requête
    file = request.files.get('file')
    if not file:
        return jsonify({"error": "No file provided"}), 400

    # Charger le fichier Excel
    try:
        df = pd.read_excel(file)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

    # Nettoyage initial des emails
    emails = df.get('Emails', [])
    if emails.empty:
        return jsonify({"error": "No 'Emails' column found in the file"}), 400

    results = []
    
    # Processus de nettoyage pour chaque email
    for email in emails.dropna():
        email_clean = clean_email(email)
        syntax_valid = is_valid_syntax(email_clean)
        domain_valid = is_domain_valid(email_clean) if syntax_valid else False
        suspect = is_suspect_email(email_clean)
        
        if syntax_valid and domain_valid and not suspect:
            statut = 'Valide'
        elif not syntax_valid:
            statut = 'Syntaxe Invalide'
        elif not domain_valid:
            statut = 'Domaine Inexistant'
        elif suspect:
            statut = 'Email Suspect'
        else:
            statut = 'Invalide'
        
        results.append({'Email': email_clean, 'Statut': statut})

    # Convertir les résultats en DataFrame
    df_results = pd.DataFrame(results)

    # Sauvegarder les résultats dans un fichier Excel
    output_path = 'cleaned_emails.xlsx'
    with pd.ExcelWriter(output_path) as writer:
        df_results.to_excel(writer, index=False)

    # Renvoyer le fichier Excel généré en réponse
    return send_file(output_path, as_attachment=True, download_name='cleaned_emails.xlsx', mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# Lancer l'application Flask
if __name__ == '__main__':
    app.run(debug=True)
