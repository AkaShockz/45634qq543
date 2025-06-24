import os
import json
import bcrypt

USERS_FILE = os.path.join(os.path.dirname(__file__), '..', 'users.json')

def load_users():
    if os.path.exists(USERS_FILE):
        with open(USERS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def save_users(users):
    with open(USERS_FILE, 'w', encoding='utf-8') as f:
        json.dump(users, f, ensure_ascii=False, indent=2)

def hash_password(password):
    return bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')

def add_user(users):
    username = input('Enter new username: ').strip()
    if username in users:
        print('User already exists.')
        return
    password = input('Enter password: ')
    users[username] = {'password': hash_password(password), 'enabled': True}
    print(f'User {username} added and enabled.')

def set_password(users):
    username = input('Enter username: ').strip()
    if username not in users:
        print('User does not exist.')
        return
    password = input('Enter new password: ')
    users[username]['password'] = hash_password(password)
    print(f'Password updated for {username}.')

def enable_user(users):
    username = input('Enter username to enable: ').strip()
    if username not in users:
        print('User does not exist.')
        return
    users[username]['enabled'] = True
    print(f'User {username} enabled.')

def disable_user(users):
    username = input('Enter username to disable: ').strip()
    if username not in users:
        print('User does not exist.')
        return
    users[username]['enabled'] = False
    print(f'User {username} disabled.')

def main():
    users = load_users()
    while True:
        print('\nUser Management:')
        print('1. Add user')
        print('2. Set password')
        print('3. Enable user')
        print('4. Disable user')
        print('5. List users')
        print('6. Exit')
        choice = input('Choose an option: ').strip()
        if choice == '1':
            add_user(users)
        elif choice == '2':
            set_password(users)
        elif choice == '3':
            enable_user(users)
        elif choice == '4':
            disable_user(users)
        elif choice == '5':
            for u, v in users.items():
                print(f"{u} - {'ENABLED' if v.get('enabled', False) else 'DISABLED'}")
        elif choice == '6':
            break
        else:
            print('Invalid choice.')
        save_users(users)

if __name__ == '__main__':
    main() 