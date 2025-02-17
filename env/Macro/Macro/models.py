from django.db import models

class GeneratedFile(models.Model):
    file_name = models.CharField(max_length=255)
    file_path = models.TextField()
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return self.file_name
