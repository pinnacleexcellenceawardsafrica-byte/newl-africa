from django.urls import path
from django.conf import settings
from django.conf.urls.static import static
from . import views

urlpatterns = [
    path('', views.index, name='index'),

    # Health check for Railway
    path('health/', views.health_check, name='health_check'),

    # API endpoints
    path('api/upload/', views.api_upload, name='api_upload'),
    path('api/sites/', views.api_sites, name='api_sites'),
    path('api/status/', views.api_status, name='api_status'),
    path('api/standings/', views.api_standings, name='api_standings'),
    path('api/import-logs/', views.api_import_logs, name='api_import_logs'),

    # Generation
    path('api/generate/<str:site_id>/', views.api_generate, name='api_generate'),
    path('api/generate-all/', views.api_generate_all, name='api_generate_all'),
    path('api/convert-to-pdf/<str:site_id>/', views.api_convert_to_pdf, name='api_convert_to_pdf'),

    # Downloads - Simplified to only Excel and signed PDF
    path('api/download/<str:cert_id>/', views.api_download, name='api_download'),
    path('api/download-pdf/<str:cert_id>/', views.api_download_pdf, name='api_download_pdf'),
    path('api/export-all/', views.api_export_all, name='api_export_all'),

    # Download sorted Excel without merging
    path('api/download-sorted-excel/', views.api_download_sorted_excel, name='api_download_sorted_excel'),

    # Download certificate as Excel only
    path('api/download-certificate-excel/<str:site_id>/', views.api_download_certificate_excel_only, name='api_download_certificate_excel_only'),

    # Download with signature (full certificate) - this is the only PDF version
    path('api/download-certificate-pdf/<str:site_id>/', views.api_download_certificate_pdf, name='api_download_certificate_pdf'),

    # File management
    path('api/update-file/', views.api_update_file, name='api_update_file'),
    path('api/delete-file/<str:file_type>/', views.api_delete_file, name='api_delete_file'),
    path('api/delete-all-files/', views.api_delete_all_files, name='api_delete_all_files'),
    path('api/clear-data/', views.api_clear_data, name='api_clear_data'),

    # Delete
    path('api/delete-certificate/<int:cert_id>/', views.api_delete_certificate, name='api_delete_certificate'),
    path('api/delete-site/<str:site_id>/', views.api_delete_site, name='api_delete_site'),
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)