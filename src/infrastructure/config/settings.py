import os
from dotenv import load_dotenv

class Settings:
  def __init__(self):
    load_dotenv()

    self.EMAIL_TEST = os.getenv("EMAIL_TEST")
    self.COMPANY = os.getenv("COMPANY")