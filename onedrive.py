import streamlit as st
import urllib.request


url = 'https://onedrive.live.com/personal/6375b5f36d64b63e/_layouts/15/Doc.aspx?resid=6375B5F36D64B63E!sc587c8a8c8c74a68b294e6f3391d5a0d&cid=6375b5f36d64b63e&migratedtospo=true&app=Excel'

urllib.request.urlretrieve(url, "Book 3.xlsx")



