import streamlit as st
import subprocess

# Layout header met logo
col1, col2 = st.columns([5, 1])

with col1:
    st.markdown(
        "<h1 style='color:#009FE3;'>Projectadministration</h1>",
        unsafe_allow_html=True
    )
with col2:
    st.image(
        "C:/Users/bram.gerrits/Desktop/Automations/ProjectAdministration/Overig/VHE logo.jpg",
        width=120
    )

# Lichte scheidingslijn
st.markdown("<hr style='border:1px solid #4F81BD'>", unsafe_allow_html=True)

# Weekly Sync
st.markdown(
    """
    <div style='background-color:#DDEBF7;padding:10px;border-radius:5px;margin-bottom:10px;'>
        <h4 style='color:#4F81BD;margin:5px;'>🔄 Weekly Sync</h4>
    </div>
    """,
    unsafe_allow_html=True
)

if st.button("🚀 Run Weekly Sync", use_container_width=True):
    with st.spinner("Running complete weekly sync..."):
        subprocess.run(
    ["python", "sync.py"],
    cwd=r"C:/Users/bram.gerrits/Desktop/Automations/ProjectAdministration/Scripts"
)
    st.success("✅ Weekly sync volledig afgerond!")

# Rapportages
st.markdown(
    """
    <div style='background-color:#DDEBF7;padding:10px;border-radius:5px;margin-top:10px;'>
        <h4 style='color:#4F81BD;margin:5px;'>📊 Rapportages</h4>
    </div>
    """,
    unsafe_allow_html=True
)

if st.button("📊 Run Rapportage", use_container_width=True):
    subprocess.run([
    "python",
    r"C:/Users/bram.gerrits/Desktop/Automations/ProjectAdministration/Scripts/mail_to_projectleaders.py"
])
    st.success("✅ Rapportage gegenereerd!")

# Mails
st.markdown(
    """
    <div style='background-color:#DDEBF7;padding:10px;border-radius:5px;margin-top:10px;'>
        <h4 style='color:#4F81BD;margin:5px;'>✉️ Mails</h4>
    </div>
    """,
    unsafe_allow_html=True
)

if st.button("✉️ Mail to Projectleaders", use_container_width=True):
    subprocess.run([
        "python",
        r"C:/Users/bram.gerrits/Desktop/Automations/ProjectAdministration/Scripts/mail_to_projectleaders.py"
    ])
    st.success("✅ Mails verstuurd!")
