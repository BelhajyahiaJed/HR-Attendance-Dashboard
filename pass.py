from streamlit_authenticator.utilities.hasher import Hasher

password = "admin"  # plain text password
hashed_password = Hasher.hash(password)

print("Hashed password:", hashed_password)
