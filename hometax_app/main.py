import os
import warnings
from gui.interface import DisplayInterface

def main():
    app = DisplayInterface()
    app.mainloop()

if __name__ == "__main__":
    os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
    warnings.filterwarnings('ignore')
    main()