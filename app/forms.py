from flask_wtf import FlaskForm
from wtforms import EmailField, PasswordField, StringField, SubmitField
from wtforms.validators import DataRequired, Email, Length, Optional


class LoginForm(FlaskForm):
    email = EmailField("Email", validators=[DataRequired(), Email()])
    password = PasswordField("Password", validators=[DataRequired()])
    submit = SubmitField("Sign In")


class RegisterForm(FlaskForm):
    full_name = StringField("Full name", validators=[DataRequired(), Length(max=120)])
    student_id = StringField("Student ID", validators=[Optional(), Length(max=32)])
    email = EmailField("Email", validators=[DataRequired(), Email(), Length(max=255)])
    password = PasswordField(
        "Password",
        validators=[DataRequired(), Length(min=8, message="Use at least 8 characters.")],
    )
    submit = SubmitField("Create Account")
