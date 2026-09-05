from django.contrib import admin
from .models import Site, PurchaseOrder, Certificate, UploadedFile, ImportLog

@admin.register(Site)
class SiteAdmin(admin.ModelAdmin):
    list_display = ['site_id', 'site_name', 'po_number', 'wcc_status', 'sub_category', 'created_at']
    list_filter = ['wcc_status', 'sub_category', 'project_category']
    search_fields = ['site_id', 'site_name', 'po_number', 'smp']
    readonly_fields = ['created_at', 'updated_at']

@admin.register(PurchaseOrder)
class PurchaseOrderAdmin(admin.ModelAdmin):
    list_display = ['po_number', 'site', 'item_no', 'description', 'quantity']
    list_filter = ['site']
    search_fields = ['po_number', 'description']

@admin.register(Certificate)
class CertificateAdmin(admin.ModelAdmin):
    list_display = ['file_name', 'site', 'format_type', 'generated_date', 'generated_by']
    list_filter = ['format_type', 'generated_date']
    search_fields = ['file_name', 'site__site_id']

@admin.register(UploadedFile)
class UploadedFileAdmin(admin.ModelAdmin):
    list_display = ['file_type', 'file_name', 'row_count', 'uploaded_date', 'is_active']
    list_filter = ['file_type', 'is_active']

@admin.register(ImportLog)
class ImportLogAdmin(admin.ModelAdmin):
    list_display = ['action', 'user', 'sites_imported', 'po_lines_imported', 'created_at']
    list_filter = ['created_at']