Install the dependencies
conda install -c conda-forge httplib2
conda install -c conda-forge google-api-python-client

Copy the secret File

"c:\Program Files\PuTTY\pscp" -i  C:\temp\mygit\akak.ppk sheets.googleapis.com-python-quickstart.json  ec2-user@ec2-52-34-181-153.us-west-2.compute.amazonaws.com:sheets.googleapis.com-python-quickstart.json
