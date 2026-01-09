# utils/identity.py
import streamlit as st

def get_executor_identity() -> str:
    """
    Retorna sempre string vazia.
    A responsabilidade de guardar o nome fica exclusivamente no app.py via st.session_state,
    garantindo que cada utilizador vê apenas o seu próprio nome.
    
    Mantemos esta função para compatibilidade com o código,
    mas ela não tenta adivinhar o utilizador no sistema.
    """
    return ""


def get_query_parameters() -> dict:
    """
    Wrapper simples para substituir o deprecated st.experimental_get_query_params.
    Atualmente não é usado para identidade, mas mantemos para compatibilidade e futura expansão.
    """
    try:
        # st.query_params já retorna um dict-like
        return dict(st.query_params)
    except Exception:
        return {}
