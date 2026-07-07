import streamlit as st
import yaml
from yaml.loader import SafeLoader
import streamlit_authenticator as stauth

st.set_page_config(page_title="Login", page_icon="👋", layout="wide")

with open("./config.yaml", "r") as file:
    config = yaml.load(file, Loader=SafeLoader)

for i, username in enumerate(config["credentials"]["usernames"]):
    # config["credentials"]["passwords"][i] = st.secrets.paswords[username]
    config["credentials"]["usernames"][username]["password"] = st.secrets.passwords[username]


def main():
    authenticator = stauth.Authenticate(
        config["credentials"],
        config["cookie"]["name"],
        config["cookie"]["key"],
        config["cookie"]["expiry_days"],
    )

    authenticator.login("main")

    if st.session_state["authentication_status"]:
        st.write(f"Welcome *{st.session_state['name']}*")
        authenticator.logout("Logout", "main")
    elif st.session_state["authentication_status"] is False:
        st.error("Username/password is incorrect")
    elif st.session_state["authentication_status"] is None:
        st.warning("Please enter your username and password")


if __name__ == "__main__":
    main()
