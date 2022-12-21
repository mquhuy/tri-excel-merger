# Tri Excel merger

This is just a rough script I wrote for my friend, whose name is Tri, hence the project name.

## Instruction

Run it like how you'd run any python project. Checkout [this link](https://learnpython.com/blog/how-to-use-virtualenv-python/) for a sample instruction to use Virtualenv.
`Python 3.5+` is required.

Instruction for Linux/MacOS:

- Create virtualenv
```
python3 -m venv venv
```
- Activate virtualenv
```
source venv/bin/activate
```
- Install required packages
```
pip install -r requirements.txt
```
- Create a `data` sub-directory
```
mkdir data
```
- Put all input excel files into the `data` sub-directory
- Run the script `python process.py`
- Result files are stored in `output` subdir.

*Note*: The script assumes that the month names are unique, so don't run data from, say, Nov 2021 and Nov 2022 in the same batch.
