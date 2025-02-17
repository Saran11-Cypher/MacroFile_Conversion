from django import forms

class FileUploadForm(forms.Form):
    excel_file = forms.FileField()
