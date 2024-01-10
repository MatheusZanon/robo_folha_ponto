"""IMPORTAÇÕES PARA AUTOMAÇÃO WEB NO CHROME COM SELENIUM"""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from time import sleep

def auto_web():
    driver = webdriver.Chrome()
    return driver, sleep, By
