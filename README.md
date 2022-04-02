# Supplier Cars data integration

a script that takes supplier_car.json and saves it into excel extemp.xlsx <br /> 
extemp.xlsx contains 3 subsheets: <br />
- preprocessing
- normalisation
- integration (containing target data under which input data must be appended)

### Installation and Requirements

_create a dedicated virtual environment using the following command_ <br /> 
python3 -m venv _yourenv_ <br /> 
source yourenv/activate

_to install from requirements file from the path where reqs.txt is located_ <br /> 
pip install -r requirements.txt 