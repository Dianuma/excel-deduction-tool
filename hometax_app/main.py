import os
import warnings
from gui.interface import display_interface

def main():
    app = display_interface()
    app.mainloop()

if __name__ == "__main__":
    os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
    warnings.filterwarnings('ignore')
    main()