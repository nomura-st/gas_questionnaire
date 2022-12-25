# gas_questionnaire

## clasp

### init

https://github.com/google/clasp

1. install

```
npm install -g @google/clasp
npm install @types/google-apps-script
```

2. Then enable the Google Apps Script API:
   https://script.google.com/home/usersettings

3. login

```
clasp login
```

4. create

```
clasp create --parentId "1D_Gxyv*****************************NXO7o"
```

- specify the doc/sheet/form/... ID
  or just `clasp create` to create with new document

5. claspignore
   add `.claspignore` to use subdirectories

6. push

```
clasp push
```

### After git clone

1. change file below

```
.clasp.json
```

- scriptId
- parentId
