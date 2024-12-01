from flask import Flask, render_template, request, redirect, url_for

app = Flask(__name__)

# Dummy database for demonstration
users = []

@app.route('/')
def home():
    # Render the welcome page
    return render_template('welcome.html')  # Assuming 'welcome.html' exists

# Route for the login page
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        email = request.form['email']
        password = request.form['password']

        # Validate login credentials
        for user in users:
            if user['email'] == email and user['password'] == password:
                return f"Welcome back, {user['name']}!"
        return "Invalid email or password. Please try again."
    return render_template('login.html')

# Route for the sign-up page
@app.route('/signup', methods=['GET', 'POST'])
def signup():
    if request.method == 'POST':
        name = request.form['name']
        email = request.form['email']
        password = request.form['password']
        confirm_password = request.form['confirm-password']

        # Validate sign-up inputs
        if password != confirm_password:
            return "Passwords do not match. Please try again."
        for user in users:
            if user['email'] == email:
                return "Email already registered. Please log in."

        # Add user to the database
        users.append({'name': name, 'email': email, 'password': password})
        return redirect(url_for('login'))
    return render_template('signup.html')

if __name__ == '__main__':
    app.run(debug=True)
