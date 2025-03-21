from app import app, db, User
import bcrypt

def init_db():
    with app.app_context():
        # Create all tables
        db.create_all()
        
        # Create default users if they don't exist
        default_users = [
            {
                'username': 'krishna',
                'password': 'krishna@123',
                'role': 'admin'
            },
            {
                'username': 'kuber',
                'password': 'kuber@123',
                'role': 'admin'
            }
        ]
        
        for user_data in default_users:
            user = User.query.filter_by(username=user_data['username']).first()
            if not user:
                # Generate a new salt and hash the password
                salt = bcrypt.gensalt()
                hashed = bcrypt.hashpw(user_data['password'].encode('utf-8'), salt)
                # Store the complete hash (including salt) as a string
                user = User(
                    username=user_data['username'],
                    password=hashed.decode('utf-8'),
                    role=user_data['role']
                )
                db.session.add(user)
        
        db.session.commit()
        print("Database initialized successfully!")

if __name__ == '__main__':
    init_db() 