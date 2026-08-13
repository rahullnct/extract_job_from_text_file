from pathlib import Path

from django import forms


class JobTextUploadForm(forms.Form):

    text_file = forms.FileField(
        label="Upload Job Text File",
        widget=forms.FileInput(
            attrs={
                "accept": ".txt",
                "class": "file-input",
            }
        ),
    )

    source_url = forms.URLField(
        label="Source Job URL",
        required=False,
        widget=forms.URLInput(
            attrs={
                "class": "text-input",
                "placeholder": "https://www.shine.com/jobs/...",
            }
        ),
    )

    def clean_text_file(self):

        uploaded_file = self.cleaned_data["text_file"]

        extension = Path(
            uploaded_file.name
        ).suffix.lower()

        if extension != ".txt":
            raise forms.ValidationError(
                "Only .txt files are allowed."
            )

        if uploaded_file.size > 5 * 1024 * 1024:
            raise forms.ValidationError(
                "Maximum file size is 5 MB."
            )

        return uploaded_file
