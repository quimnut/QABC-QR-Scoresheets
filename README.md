# QABC-QR-Scoresheets
Scoresheet tool for brewing competitions using BCOE&M

Takes the entries paid csv export to generate scoresheets with judging details and QR codes for post processing.

Post processes scoresheets into judging numbers for upload. Process bad scans and manually assign.

## Usage
Under windows I suggest WSL/Ubuntu and an x11 server, say mobaxterm or something.
```
git clone git@github.com:quimnut/QABC-QR-Scoresheets.git
cd QABC-QR-Scoresheets/
python3 -m venv .
source bin/activate
pip3 install --upgrade pip
pip3 install -r requirements.txt
python main.py
```

## Usage
Select the scoresheets you choose to use, the CSV (with tables assigned).

![generate](https://user-images.githubusercontent.com/2128947/141775701-4aece719-ef86-47cd-aa8c-4ba85b26de2b.png)

Preview a PDF to see the fields look ok etc.

![previewsheet](https://user-images.githubusercontent.com/2128947/141775725-32225290-f18b-4eb7-8ba1-eb32f6a7f7e1.png)

Preview the CSV data to check tables and flights look ok

![previewentries](https://user-images.githubusercontent.com/2128947/141775723-31ab1532-e58e-4d9d-95e2-d199e72b2be3.png)

Generate your PDFs!

After the competition scan the sheets to pdf. Select your PDF and the destination folder

![process](https://user-images.githubusercontent.com/2128947/141775718-d29aa386-38ce-4fae-9617-bd0d8a4caa8c.png)

If there are failed scans, choose the page to review. 

![selectreject](https://user-images.githubusercontent.com/2128947/141775739-ea05000e-5032-4608-a252-9b4de0881b43.png)

You will have the option to delete the page, or assign a judging number, in which case it will copy the page and delete it from the rejects file.

![processreject](https://user-images.githubusercontent.com/2128947/141775720-e2b25125-de9b-47ae-bdd6-ed6a89d1c104.png)

# Windows binary
[QABC Scoresheet Wizard v0.1.zip](https://objectstorage.ap-sydney-1.oraclecloud.com/p/Un0o3uRORVQk6U35sz40OtBQ6CQH-8_ex1womhe0wkPQAQEU_0mAqsp6vzsYvuVK/n/sdodooe6mbu1/b/goat-bucket/o/QABC%20Scoresheet%20Wizard%20v0.1.zip)
