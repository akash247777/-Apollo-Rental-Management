from app import app, DB_SERVER, DB_DATABASE, DB_USER

if __name__ == '__main__':
    print(f"Connecting to database {DB_DATABASE} on server {DB_SERVER} as user {DB_USER}")
    app.run(debug=True, port=5000) 