from django.urls import path

from . import views


app_name = "extractor"


urlpatterns = [
    path(
        "",
        views.upload_job_text,
        name="upload_job_text",
    ),
]

