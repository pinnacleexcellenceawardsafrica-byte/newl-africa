from django.db import models
from django.contrib.auth.models import User

class UploadedFile(models.Model):
    FILE_TYPES = (
        ('book2', 'Book2 Master File'),
        ('purchase_orders', 'Purchase Orders'),
        ('template', 'Certificate Template'),
    )
    
    file_type = models.CharField(max_length=20, choices=FILE_TYPES)
    file_name = models.CharField(max_length=255)
    file_path = models.FileField(upload_to='uploads/%Y/%m/%d/')
    row_count = models.IntegerField(default=0)
    uploaded_date = models.DateTimeField(auto_now_add=True)
    is_active = models.BooleanField(default=True)
    uploaded_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    
    def __str__(self):
        return f"{self.get_file_type_display()} - {self.file_name}"


class Site(models.Model):
    STATUS_CHOICES = (
        ('Pending', 'Pending'),
        ('Done', 'Done'),
    )
    
    no = models.CharField(max_length=10, blank=True)
    site_id = models.CharField(max_length=50)
    site_name = models.CharField(max_length=200)
    smp = models.CharField(max_length=100, blank=True, db_index=True, unique=True)
    site_module_package = models.CharField(max_length=200, blank=True)  # NEW: Used for mapping
    po_number = models.CharField(max_length=20, db_index=True)
    po_date = models.CharField(max_length=20, blank=True)
    wcc_status = models.CharField(max_length=20, choices=STATUS_CHOICES, default='Pending')
    nokia_pm = models.CharField(max_length=100, blank=True, default='Julius Kamemba')
    sub_category = models.CharField(max_length=50, blank=True)
    project_category = models.CharField(max_length=50, blank=True)
    region = models.CharField(max_length=100, blank=True)
    remarks = models.TextField(blank=True)
    
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    source_file = models.ForeignKey(UploadedFile, on_delete=models.SET_NULL, null=True, blank=True)
    
    def __str__(self):
        return f"{self.smp} - {self.site_name} (PO: {self.po_number})"
    
    def get_filename(self):
        """
        Certificate / export filename.

        Strictly "Site ID-Site Name" — SMP is never used for naming,
        even when present. This matches the "Site ID & Site Name" field
        shown on the certificate itself.
        """
        safe_id = (self.site_id or '').replace('/', '_').replace('\\', '_').replace(':', '_').strip()
        safe_name = (self.site_name or '').replace('/', '_').replace('\\', '_').replace(':', '_').strip()
        return f"{safe_id}-{safe_name}"
    
    class Meta:
        ordering = ['smp']


class PurchaseOrder(models.Model):
    site = models.ForeignKey(Site, on_delete=models.CASCADE, related_name='po_items')
    po_number = models.CharField(max_length=20, db_index=True)
    po_date = models.CharField(max_length=20, blank=True)
    item_no = models.CharField(max_length=20)
    unit = models.CharField(max_length=10, default='PCE')
    quantity = models.IntegerField(default=1)
    description = models.TextField()
    remarks = models.TextField(blank=True, null=True)
    
    created_at = models.DateTimeField(auto_now_add=True)
    source_file = models.ForeignKey(UploadedFile, on_delete=models.SET_NULL, null=True, blank=True)
    
    def __str__(self):
        return f"{self.po_number} - {self.description[:50]} (SMP: {self.site.smp})"
    
    class Meta:
        ordering = ['site', 'item_no']
        unique_together = ['site', 'po_number', 'item_no', 'description']


class Certificate(models.Model):
    FORMAT_CHOICES = (
        ('excel', 'Excel'),
        ('pdf', 'PDF'),
        ('both', 'Both'),
    )
    
    site = models.ForeignKey(Site, on_delete=models.CASCADE, related_name='certificates')
    file_name = models.CharField(max_length=255)
    file_path = models.FileField(upload_to='certificates/%Y/%m/%d/')
    pdf_file_path = models.FileField(upload_to='certificates/%Y/%m/%d/pdfs/', blank=True, null=True)
    sheet_used = models.CharField(max_length=50, blank=True)
    format_type = models.CharField(max_length=10, choices=FORMAT_CHOICES, default='both')
    generated_date = models.DateTimeField(auto_now_add=True)
    generated_by = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    is_active = models.BooleanField(default=True)
    
    def __str__(self):
        return f"Certificate: {self.file_name}"
    
    class Meta:
        ordering = ['-generated_date']


class ImportLog(models.Model):
    user = models.ForeignKey(User, on_delete=models.SET_NULL, null=True, blank=True)
    action = models.CharField(max_length=100)
    details = models.TextField(blank=True)
    sites_imported = models.IntegerField(default=0)
    po_lines_imported = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return f"{self.action} - {self.created_at.strftime('%Y-%m-%d %H:%M')}"
    
    class Meta:
        ordering = ['-created_at']