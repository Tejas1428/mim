import requests

url = "https://dev.azure.com/daimler-mic/devdataproc/_git/meta-info-models?path=/dte"
try:
    resp = requests.get(url)
    print(resp.text)
except:
    print('error')