import os
import json
import zipfile
from io import BytesIO
from datetime import datetime
from django.shortcuts import render, get_object_or_404
from django.http import JsonResponse, FileResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt
from django.conf import settings
from django.core.files.storage import default_storage
from django.db import transaction
import traceback
import time
import shutil

from .models import Site, PurchaseOrder, Certificate, UploadedFile, ImportLog
from .excel_handler import (
    process_book2, process_purchase_orders, generate_certificate_excel,
    generate_sorted_excel, update_book2_status, update_book2_status_bulk,
    get_sheet_from_module, get_pm_for_site, get_pm_from_sheet, convert_excel_to_pdf,
    convert_excel_batch_to_pdf_libreoffice, kill_excel_processes
)
from .certificate import generate_certificate_pdf


def index(request):
    """Main page - serves your index.html"""
    context = {
        'sites': Site.objects.all().order_by('smp'),
        'certificates': Certificate.objects.all().order_by('-generated_date')[:20],
        'total_sites': Site.objects.count(),
        'pending_sites': Site.objects.filter(wcc_status='Pending').count(),
        'done_sites': Site.objects.filter(wcc_status='Done').count(),
        'total_certificates': Certificate.objects.count(),
        'has_book2': UploadedFile.objects.filter(file_type='book2', is_active=True).exists(),
        'has_pos': UploadedFile.objects.filter(file_type='purchase_orders', is_active=True).exists(),
        'has_template': UploadedFile.objects.filter(file_type='template', is_active=True).exists(),
        'import_logs': ImportLog.objects.all()[:10],
    }
    return render(request, 'core/index.html', context)


@csrf_exempt
def api_upload(request):
    """Handle file uploads from frontend"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        with transaction.atomic():
            book2_file = request.FILES.get('book2_file')
            if not book2_file:
                return JsonResponse({'error': 'Book2 file required'}, status=400)

            book2_path = default_storage.save(f'uploads/book2/{book2_file.name}', book2_file)
            uploaded_book2 = UploadedFile.objects.create(
                file_type='book2',
                file_name=book2_file.name,
                file_path=book2_path,
                uploaded_by=request.user if request.user.is_authenticated else None
            )

            po_file = request.FILES.get('purchase_orders_file')
            if not po_file:
                return JsonResponse({'error': 'Purchase Orders file required'}, status=400)

            po_path = default_storage.save(f'uploads/purchase_orders/{po_file.name}', po_file)
            uploaded_po = UploadedFile.objects.create(
                file_type='purchase_orders',
                file_name=po_file.name,
                file_path=po_path,
                uploaded_by=request.user if request.user.is_authenticated else None
            )

            template_file = request.FILES.get('template_file')
            if not template_file:
                return JsonResponse({'error': 'Template file required'}, status=400)

            template_path = default_storage.save(f'uploads/template/{template_file.name}', template_file)
            uploaded_template = UploadedFile.objects.create(
                file_type='template',
                file_name=template_file.name,
                file_path=template_path,
                uploaded_by=request.user if request.user.is_authenticated else None
            )

            book2_abs_path = os.path.join(settings.MEDIA_ROOT, book2_path)
            sites_imported = process_book2(book2_abs_path, uploaded_book2)

            po_abs_path = os.path.join(settings.MEDIA_ROOT, po_path)
            po_imported = process_purchase_orders(po_abs_path, uploaded_po)

            request.session['template_path'] = template_path
            request.session['book2_path'] = book2_path

            ImportLog.objects.create(
                user=request.user if request.user.is_authenticated else None,
                action='Imported Files',
                details=f'Book2: {book2_file.name}, POs: {po_file.name}, Template: {template_file.name}',
                sites_imported=sites_imported,
                po_lines_imported=po_imported
            )

            return JsonResponse({
                'success': True,
                'message': f'Imported {sites_imported} sites and {po_imported} PO lines',
                'sites_imported': sites_imported,
                'po_lines_imported': po_imported,
                'total_sites': Site.objects.count(),
                'total_pos': PurchaseOrder.objects.count(),
            })

    except Exception as e:
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)


def api_sites(request):
    """Return sites data for frontend table with sorting and filtering"""
    sites = Site.objects.all().order_by('smp', 'nokia_pm', 'site_module_package')

    data = []
    for site in sites:
        cert = site.certificates.first()
        has_pdf = False
        if cert and cert.pdf_file_path:
            pdf_path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
            has_pdf = os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000

        po_items = PurchaseOrder.objects.filter(site=site)

        data.append({
            'id': site.id,
            'smp': site.smp,
            'site_id': site.site_id,
            'site_name': site.site_name,
            'site_module_package': site.site_module_package,
            'po_number': site.po_number,
            'po_date': site.po_date,
            'wcc_status': site.wcc_status,
            'nokia_pm': site.nokia_pm,
            'sub_category': site.sub_category,
            'project_category': site.project_category,
            'region': site.region,
            'has_certificate': bool(cert),
            'has_pdf': has_pdf,
            'certificate_id': cert.id if cert else None,
            'certificate_file': cert.file_name if cert else None,
            'po_items_count': po_items.count(),
            'po_items': [{'item_no': p.item_no, 'description': p.description, 'quantity': p.quantity, 'unit': p.unit} for p in po_items],
        })

    return JsonResponse({'sites': data, 'total': len(data)})


def health_check(request):
    """Simple health check endpoint for Railway"""
    import json
    from django.db import connection
    
    try:
        # Test database connection
        with connection.cursor() as cursor:
            cursor.execute("SELECT 1")
            db_connected = True
    except Exception as e:
        db_connected = False
        db_error = str(e)
    
    return JsonResponse({
        'status': 'healthy' if db_connected else 'degraded',
        'message': 'Certificate Generator is running',
        'database': 'connected' if db_connected else f'error: {db_error}',
        'migrations': 'applied',
        'sites_count': Site.objects.count(),
        'certificates_count': Certificate.objects.count(),
    })



def _delete_certificates_for_site(site):
    """
    Wipe every existing Certificate row (and its Excel/PDF files) for a site.
    This prevents stale files from surviving under deterministic filenames.
    """
    for old_cert in site.certificates.all():
        try:
            if old_cert.file_path:
                path = os.path.join(settings.MEDIA_ROOT, old_cert.file_path.name)
                if os.path.exists(path):
                    os.remove(path)
                    print(f"✅ Deleted old Excel: {path}")
        except Exception as e:
            print(f"⚠️ Could not delete old Excel for {site.smp}: {e}")
        try:
            if old_cert.pdf_file_path:
                path = os.path.join(settings.MEDIA_ROOT, old_cert.pdf_file_path.name)
                if os.path.exists(path):
                    os.remove(path)
                    print(f"✅ Deleted old PDF: {path}")
        except Exception as e:
            print(f"⚠️ Could not delete old PDF for {site.smp}: {e}")
        old_cert.delete()
        print(f"🗑️ Deleted old certificate record for {site.smp}")


@csrf_exempt
def api_generate(request, site_id):
    """Generate certificate for a specific site"""
    if request.session.get('generating', False):
        return JsonResponse({'error': 'Already generating certificates'}, status=400)

    request.session['generating'] = True

    try:
        site = Site.objects.filter(smp=site_id).first()
        if not site:
            site = Site.objects.filter(site_id=site_id).first()
        if not site:
            try:
                site = Site.objects.get(id=int(site_id))
            except (ValueError, Site.DoesNotExist):
                pass

        if not site:
            return JsonResponse({'error': f'Site not found: {site_id}'}, status=404)

        print(f"🔍 Generating for site: SMP={site.smp} - {site.site_name}")
        print(f"📋 PO Number: {site.po_number}")
        print(f"📂 Site Module Package: {site.site_module_package}")

        template_file = UploadedFile.objects.filter(file_type='template', is_active=True).first()
        if not template_file:
            return JsonResponse({'error': 'Please upload a certificate template first'}, status=400)

        template_path = os.path.join(settings.MEDIA_ROOT, template_file.file_path.name)

        if not os.path.exists(template_path):
            return JsonResponse({'error': f'Template file not found'}, status=400)

        po_items = PurchaseOrder.objects.filter(site=site)
        print(f"📦 PO items found for PO {site.po_number}: {po_items.count()}")

        if not po_items.exists():
            return JsonResponse({'error': f'No PO items found for SMP {site.smp}'}, status=400)

        excel_dir = os.path.join(settings.MEDIA_ROOT, 'certificates', 'excel')
        pdf_dir = os.path.join(settings.MEDIA_ROOT, 'certificates', 'pdf')
        os.makedirs(excel_dir, exist_ok=True)
        os.makedirs(pdf_dir, exist_ok=True)

        filename = site.get_filename()
        excel_filename = f"{filename}.xlsx"
        excel_path = os.path.join(excel_dir, excel_filename)

        # Wipe any previous certificate for this site FIRST
        _delete_certificates_for_site(site)

        # Single-site generation converts its own PDF immediately
        result = generate_certificate_excel(site, po_items, template_path, excel_path)

        if isinstance(result, tuple):
            excel_path, signed_pdf_path = result
        else:
            signed_pdf_path = None

        if not os.path.exists(excel_path):
            return JsonResponse({'error': 'Excel file was not created'}, status=500)

        signed_pdf_exists = signed_pdf_path and os.path.exists(signed_pdf_path) and os.path.getsize(signed_pdf_path) > 1000

        cert = Certificate.objects.create(
            site=site,
            file_name=excel_filename,
            file_path=f'certificates/excel/{excel_filename}',
            pdf_file_path=f'certificates/pdf/{os.path.basename(signed_pdf_path)}' if signed_pdf_exists else None,
            sheet_used=site.site_module_package or 'Modernization',
            format_type='both' if signed_pdf_exists else 'excel_only',
            generated_by=request.user if request.user.is_authenticated else None
        )

        site.wcc_status = 'Done'
        site.save()

        book2_path = request.session.get('book2_path')
        if book2_path:
            book2_abs_path = os.path.join(settings.MEDIA_ROOT, book2_path)
            for attempt in range(3):
                try:
                    update_book2_status(book2_abs_path, site.smp, 'Completed')
                    break
                except Exception as e:
                    print(f"⚠️ Book2 update attempt {attempt+1} failed: {e}")
                    kill_excel_processes()
                    time.sleep(1)

        ImportLog.objects.create(
            user=request.user if request.user.is_authenticated else None,
            action='Generated Certificate',
            details=f'Generated certificate for SMP {site.smp} - {site.site_name} (PO: {site.po_number})',
            sites_imported=1,
            po_lines_imported=po_items.count()
        )

        return JsonResponse({
            'success': True,
            'message': f'Certificate generated for {site.site_name}',
            'certificate_id': cert.id,
            'excel_url': f'/media/certificates/excel/{excel_filename}',
            'pdf_url': f'/media/certificates/pdf/{os.path.basename(signed_pdf_path)}' if signed_pdf_exists else None,
            'file_name': excel_filename,
            'has_pdf': signed_pdf_exists,
            'total_sites': Site.objects.count(),
            'pending_sites': Site.objects.filter(wcc_status='Pending').count(),
            'done_sites': Site.objects.filter(wcc_status='Done').count(),
            'total_certificates': Certificate.objects.count(),
        })

    except Exception as e:
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)
    finally:
        request.session['generating'] = False


@csrf_exempt
def api_generate_all(request):
    """
    Generate certificates for all sites — batched for speed.
    """
    if request.session.get('generating', False):
        return JsonResponse({'error': 'Already generating certificates'}, status=400)

    request.session['generating'] = True

    try:
        all_sites = Site.objects.all()

        print(f"🔍 Total sites: {all_sites.count()}")

        if not all_sites.exists():
            return JsonResponse({
                'message': 'No sites found. Please import Book2 first.',
                'success_count': 0
            })

        template_file = UploadedFile.objects.filter(file_type='template', is_active=True).first()
        if not template_file:
            return JsonResponse({
                'error': 'Please upload a certificate template first',
                'success_count': 0
            }, status=400)

        template_path = os.path.join(settings.MEDIA_ROOT, template_file.file_path.name)

        if not os.path.exists(template_path):
            return JsonResponse({
                'error': 'Template file not found',
                'success_count': 0
            }, status=400)

        excel_dir = os.path.join(settings.MEDIA_ROOT, 'certificates', 'excel')
        pdf_dir = os.path.join(settings.MEDIA_ROOT, 'certificates', 'pdf')
        os.makedirs(excel_dir, exist_ok=True)
        os.makedirs(pdf_dir, exist_ok=True)

        total = all_sites.count()
        error_count = 0
        errors = []

        # --- Pass 1: generate every site's Excel ---
        generated = {}

        for idx, site in enumerate(all_sites):
            try:
                print(f"🔄 Generating Excel {idx+1}/{total}: SMP={site.smp} - {site.site_name}")

                po_items = PurchaseOrder.objects.filter(site=site)
                print(f"   📦 PO items found for PO {site.po_number}: {po_items.count()}")

                if not po_items.exists():
                    error_count += 1
                    errors.append(f"{site.smp}: No PO items found")
                    continue

                filename = site.get_filename()
                excel_filename = f"{filename}.xlsx"
                excel_path = os.path.join(excel_dir, excel_filename)

                # Wipe any previous certificate for this site before writing
                _delete_certificates_for_site(site)

                result = generate_certificate_excel(
                    site, po_items, template_path, excel_path, convert_pdf=False
                )
                excel_path = result[0] if isinstance(result, tuple) else excel_path

                if not os.path.exists(excel_path):
                    error_count += 1
                    errors.append(f"{site.smp}: Excel file not created")
                    continue

                generated[site.id] = {
                    'site': site,
                    'excel_path': excel_path,
                    'excel_filename': excel_filename,
                }

            except Exception as e:
                error_count += 1
                errors.append(f"{site.smp}: {str(e)}")
                print(f"   ❌ Error generating {site.smp}: {str(e)}")
                traceback.print_exc()

        # --- Pass 2: batch-convert every generated Excel file to PDF at once ---
        excel_paths = [g['excel_path'] for g in generated.values()]
        pdf_results = convert_excel_batch_to_pdf_libreoffice(excel_paths, pdf_dir)

        # --- Pass 3: create Certificate records, flip status, collect SMPs for Book2 ---
        success_count = 0
        certificates_generated = []
        completed_smps = []

        for site_id, g in generated.items():
            site = g['site']
            excel_filename = g['excel_filename']
            base_name = os.path.splitext(excel_filename)[0]
            pdf_path = os.path.join(pdf_dir, f"{base_name}.pdf")
            signed_pdf_exists = pdf_results.get(g['excel_path'], False) and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000

            cert = Certificate.objects.create(
                site=site,
                file_name=excel_filename,
                file_path=f'certificates/excel/{excel_filename}',
                pdf_file_path=f'certificates/pdf/{base_name}.pdf' if signed_pdf_exists else None,
                sheet_used=site.site_module_package or 'Modernization',
                format_type='both' if signed_pdf_exists else 'excel_only',
                generated_by=request.user if request.user.is_authenticated else None
            )

            certificates_generated.append(cert.id)
            site.wcc_status = 'Done'
            site.save()
            completed_smps.append(site.smp)
            success_count += 1
            print(f"   ✅ Certificate {success_count}/{total} generated successfully (PDF: {'yes' if signed_pdf_exists else 'pending'})")

        # --- Pass 4: one Book2 write for every completed SMP ---
        book2_path = request.session.get('book2_path')
        if book2_path and completed_smps:
            book2_abs_path = os.path.join(settings.MEDIA_ROOT, book2_path)
            for attempt in range(3):
                try:
                    if update_book2_status_bulk(book2_abs_path, completed_smps, 'Completed'):
                        break
                except Exception as e:
                    print(f"⚠️ Book2 bulk update attempt {attempt+1} failed: {e}")
                    kill_excel_processes()
                    time.sleep(1)

        ImportLog.objects.create(
            user=request.user if request.user.is_authenticated else None,
            action='Generated All Certificates',
            details=f'Generated {success_count} certificates, {error_count} errors',
            sites_imported=success_count,
            po_lines_imported=0
        )

        return JsonResponse({
            'success': True,
            'message': f'Generated {success_count} certificates. Errors: {error_count}',
            'success_count': success_count,
            'error_count': error_count,
            'errors': errors,
            'certificate_ids': certificates_generated,
            'total_sites': Site.objects.count(),
            'pending_sites': Site.objects.filter(wcc_status='Pending').count(),
            'done_sites': Site.objects.filter(wcc_status='Done').count(),
            'total_certificates': Certificate.objects.count(),
        })

    except Exception as e:
        print(f"❌ Fatal error: {str(e)}")
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)
    finally:
        request.session['generating'] = False


def api_download_sorted_excel(request):
    """Download a sorted Excel file with all sites and PO items (no merging)"""
    try:
        sites = Site.objects.all().order_by('nokia_pm', 'site_module_package', 'site_name')

        if not sites.exists():
            return JsonResponse({'error': 'No sites found'}, status=400)

        excel_dir = os.path.join(settings.MEDIA_ROOT, 'exports')
        os.makedirs(excel_dir, exist_ok=True)

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"sites_export_{timestamp}.xlsx"
        file_path = os.path.join(excel_dir, filename)

        generate_sorted_excel(sites, file_path)

        if os.path.exists(file_path):
            response = FileResponse(open(file_path, 'rb'),
                                   content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = f'attachment; filename="{filename}"'
            return response

        return JsonResponse({'error': 'Failed to generate export'}, status=500)

    except Exception as e:
        print(f"❌ Error exporting sorted Excel: {e}")
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)


def api_download_certificate_excel_only(request, site_id):
    """Download certificate as Excel only (without PDF)"""
    try:
        site = Site.objects.filter(smp=site_id).first()
        if not site:
            site = Site.objects.filter(site_id=site_id).first()
        if not site:
            try:
                site = Site.objects.get(id=int(site_id))
            except (ValueError, Site.DoesNotExist):
                pass

        if not site:
            return JsonResponse({'error': f'Site not found: {site_id}'}, status=404)

        cert = site.certificates.first()
        if not cert:
            return JsonResponse({'error': 'Certificate not found'}, status=404)

        file_path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
        if os.path.exists(file_path):
            response = FileResponse(open(file_path, 'rb'),
                                   content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = f'attachment; filename="{cert.file_name}"'
            return response

        return JsonResponse({'error': 'File not found'}, status=404)

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


def api_download_certificate_pdf(request, site_id):
    """Download certificate PDF with signatures"""
    try:
        site = Site.objects.filter(smp=site_id).first()
        if not site:
            site = Site.objects.filter(site_id=site_id).first()
        if not site:
            try:
                site = Site.objects.get(id=int(site_id))
            except (ValueError, Site.DoesNotExist):
                pass

        if not site:
            return JsonResponse({'error': f'Site not found: {site_id}'}, status=404)

        cert = site.certificates.first()
        if not cert:
            return JsonResponse({'error': 'Certificate not found'}, status=404)

        if cert.pdf_file_path:
            pdf_path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
                response = FileResponse(open(pdf_path, 'rb'), content_type='application/pdf')
                response['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_path)}"'
                return response

        excel_path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
        if os.path.exists(excel_path):
            pdf_dir = os.path.join(settings.MEDIA_ROOT, 'certificates', 'pdf')
            os.makedirs(pdf_dir, exist_ok=True)

            base_name = cert.file_name.replace('.xlsx', '')
            pdf_path = os.path.join(pdf_dir, f"{base_name}.pdf")

            success = convert_excel_to_pdf(excel_path, pdf_path)
            if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
                cert.pdf_file_path = f'certificates/pdf/{base_name}.pdf'
                cert.format_type = 'both'
                cert.save()
                response = FileResponse(open(pdf_path, 'rb'), content_type='application/pdf')
                response['Content-Disposition'] = f'attachment; filename="{base_name}.pdf"'
                return response

        return JsonResponse({'error': 'PDF not found'}, status=404)

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@csrf_exempt
def api_convert_to_pdf(request, site_id):
    """Convert an existing Excel certificate to PDF."""
    try:
        site = Site.objects.filter(smp=site_id).first()
        if not site:
            site = Site.objects.filter(site_id=site_id).first()
        if not site:
            try:
                site = Site.objects.get(id=int(site_id))
            except (ValueError, Site.DoesNotExist):
                pass

        if not site:
            return JsonResponse({'success': False, 'error': f'Site not found: {site_id}'})

        cert = site.certificates.first()
        if not cert:
            return JsonResponse({'success': False, 'error': 'Certificate not found'})

        excel_path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
        if not os.path.exists(excel_path):
            return JsonResponse({'success': False, 'error': 'Excel file not found'})

        pdf_dir = os.path.join(settings.MEDIA_ROOT, 'certificates', 'pdf')
        os.makedirs(pdf_dir, exist_ok=True)

        base_name = cert.file_name.replace('.xlsx', '')
        pdf_filename = f"{base_name}.pdf"
        pdf_path = os.path.join(pdf_dir, pdf_filename)

        # Always regenerate the PDF - remove stale file first
        if os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
                print(f"🗑️ Removed stale PDF: {pdf_path}")
            except Exception as e:
                print(f"⚠️ Could not remove stale PDF: {e}")

        kill_excel_processes()
        time.sleep(1)

        success = convert_excel_to_pdf(excel_path, pdf_path)

        if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
            cert.pdf_file_path = f'certificates/pdf/{pdf_filename}'
            cert.format_type = 'both'
            cert.save()

            return JsonResponse({
                'success': True,
                'message': 'PDF converted successfully!',
                'pdf_url': f'/media/certificates/pdf/{pdf_filename}',
                'certificate_id': cert.id
            })
        else:
            return JsonResponse({
                'success': False,
                'error': 'PDF conversion failed. Please download the Excel file.'
            })

    except Exception as e:
        return JsonResponse({'success': False, 'error': str(e)})


def get_certificate(cert_id):
    """Helper to get certificate by ID or site ID"""
    try:
        return Certificate.objects.get(id=int(cert_id))
    except (ValueError, Certificate.DoesNotExist):
        site = Site.objects.filter(smp=cert_id).first()
        if site:
            return site.certificates.first()
        site = Site.objects.filter(site_id=cert_id).first()
        if site:
            return site.certificates.first()
        try:
            site = Site.objects.get(id=int(cert_id))
            if site:
                return site.certificates.first()
        except (ValueError, Site.DoesNotExist):
            pass
    return None


def api_download(request, cert_id):
    """Download Excel certificate"""
    cert = get_certificate(cert_id)
    if not cert:
        return JsonResponse({'error': 'Certificate not found'}, status=404)

    file_path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
    if os.path.exists(file_path):
        response = FileResponse(open(file_path, 'rb'), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename="{cert.file_name}"'
        return response

    return JsonResponse({'error': 'File not found'}, status=404)


def api_download_pdf(request, cert_id):
    """Download PDF certificate"""
    cert = get_certificate(cert_id)
    if not cert:
        return JsonResponse({'error': 'Certificate not found'}, status=404)

    if cert.pdf_file_path:
        pdf_path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
        if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
            response = FileResponse(open(pdf_path, 'rb'), content_type='application/pdf')
            response['Content-Disposition'] = f'attachment; filename="{os.path.basename(pdf_path)}"'
            return response

    excel_path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
    if os.path.exists(excel_path):
        pdf_dir = os.path.join(settings.MEDIA_ROOT, 'certificates', 'pdf')
        os.makedirs(pdf_dir, exist_ok=True)

        base_name = cert.file_name.replace('.xlsx', '')
        pdf_path = os.path.join(pdf_dir, f"{base_name}.pdf")

        success = convert_excel_to_pdf(excel_path, pdf_path)
        if success and os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 1000:
            cert.pdf_file_path = f'certificates/pdf/{base_name}.pdf'
            cert.format_type = 'both'
            cert.save()
            response = FileResponse(open(pdf_path, 'rb'), content_type='application/pdf')
            response['Content-Disposition'] = f'attachment; filename="{base_name}.pdf"'
            return response

    return JsonResponse({'error': 'PDF not found'}, status=404)


def api_delete_certificate(request, cert_id):
    """Delete a certificate"""
    cert = get_object_or_404(Certificate, id=cert_id)

    try:
        if cert.file_path:
            path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
            if os.path.exists(path):
                os.remove(path)
    except Exception as e:
        print(f"⚠️ Could not delete file: {e}")

    try:
        if cert.pdf_file_path:
            path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
            if os.path.exists(path):
                os.remove(path)
    except Exception as e:
        print(f"⚠️ Could not delete PDF file: {e}")

    site = cert.site
    cert.delete()

    if not site.certificates.exists():
        site.wcc_status = 'Pending'
        site.save()

    return JsonResponse({'success': True, 'message': 'Certificate deleted'})


def api_delete_site(request, site_id):
    """Delete a site and all its data"""
    site = Site.objects.filter(smp=site_id).first()
    if not site:
        site = Site.objects.filter(site_id=site_id).first()
    if not site:
        try:
            site = Site.objects.get(id=int(site_id))
        except (ValueError, Site.DoesNotExist):
            pass

    if not site:
        return JsonResponse({'error': f'Site not found: {site_id}'}, status=404)

    for cert in site.certificates.all():
        try:
            if cert.file_path:
                path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
                if os.path.exists(path):
                    os.remove(path)
        except:
            pass
        try:
            if cert.pdf_file_path:
                path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
                if os.path.exists(path):
                    os.remove(path)
        except:
            pass
        cert.delete()

    site.po_items.all().delete()
    site.delete()
    return JsonResponse({'success': True, 'message': f'Site {site_id} deleted'})


def api_export_all(request):
    """Export all certificates as ZIP"""
    certs = Certificate.objects.filter(is_active=True)
    if not certs.exists():
        return JsonResponse({'error': 'No certificates to export'}, status=400)

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
        for cert in certs:
            try:
                file_path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
                if os.path.exists(file_path):
                    zip_file.write(file_path, cert.file_name)
            except:
                pass

            if cert.pdf_file_path:
                try:
                    pdf_path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
                    if os.path.exists(pdf_path):
                        pdf_name = cert.file_name.replace('.xlsx', '.pdf')
                        zip_file.write(pdf_path, pdf_name)
                except:
                    pass

    response = HttpResponse(zip_buffer.getvalue(), content_type='application/zip')
    response['Content-Disposition'] = f'attachment; filename="certificates_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip"'
    return response


@csrf_exempt
def api_clear_data(request):
    """Clear all data"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        kill_excel_processes()

        for cert in Certificate.objects.all():
            try:
                if cert.file_path:
                    path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
                    if os.path.exists(path):
                        os.remove(path)
            except Exception as e:
                print(f"⚠️ Could not delete file: {e}")

            try:
                if cert.pdf_file_path:
                    path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
                    if os.path.exists(path):
                        os.remove(path)
            except Exception as e:
                print(f"⚠️ Could not delete PDF file: {e}")

        Certificate.objects.all().delete()
        PurchaseOrder.objects.all().delete()
        Site.objects.all().delete()

        for file_obj in UploadedFile.objects.all():
            try:
                if file_obj.file_path:
                    path = os.path.join(settings.MEDIA_ROOT, file_obj.file_path.name)
                    if os.path.exists(path):
                        os.remove(path)
            except Exception as e:
                print(f"⚠️ Could not delete uploaded file: {e}")
            file_obj.delete()

        request.session.pop('template_path', None)
        request.session.pop('book2_path', None)
        request.session.pop('generating', None)

        ImportLog.objects.create(
            user=request.user if request.user.is_authenticated else None,
            action='Cleared All Data',
            details='All data and files cleared',
            sites_imported=0,
            po_lines_imported=0
        )

        return JsonResponse({'success': True, 'message': 'All data cleared'})

    except Exception as e:
        print(f"❌ Error in clear_data: {str(e)}")
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)


def api_status(request):
    """Get system status"""
    book2_file = UploadedFile.objects.filter(file_type='book2', is_active=True).first()
    po_file = UploadedFile.objects.filter(file_type='purchase_orders', is_active=True).first()
    template_file = UploadedFile.objects.filter(file_type='template', is_active=True).first()

    return JsonResponse({
        'total_sites': Site.objects.count(),
        'pending_sites': Site.objects.filter(wcc_status='Pending').count(),
        'done_sites': Site.objects.filter(wcc_status='Done').count(),
        'total_certificates': Certificate.objects.count(),
        'total_po_lines': PurchaseOrder.objects.count(),
        'has_book2': bool(book2_file),
        'has_pos': bool(po_file),
        'has_template': bool(template_file),
        'book2_date': book2_file.uploaded_date.strftime('%Y-%m-%d %H:%M') if book2_file else None,
        'po_date': po_file.uploaded_date.strftime('%Y-%m-%d %H:%M') if po_file else None,
        'template_date': template_file.uploaded_date.strftime('%Y-%m-%d %H:%M') if template_file else None,
        'book2_name': book2_file.file_name if book2_file else None,
        'po_name': po_file.file_name if po_file else None,
        'template_name': template_file.file_name if template_file else None,
        'is_generating': request.session.get('generating', False),
    })


def api_standings(request):
    """Standings endpoint"""
    sites = Site.objects.all().order_by('smp')

    data = []
    total_by_po = {}

    for site in sites:
        po_items = PurchaseOrder.objects.filter(site=site)
        cert = site.certificates.first()

        if site.po_number not in total_by_po:
            total_by_po[site.po_number] = 0
        total_by_po[site.po_number] += 1

        data.append({
            'smp': site.smp,
            'site_id': site.site_id,
            'site_name': site.site_name,
            'site_module_package': site.site_module_package,
            'po_number': site.po_number,
            'po_items_count': po_items.count(),
            'wcc_status': site.wcc_status,
            'has_certificate': bool(cert),
            'certificate_id': cert.id if cert else None,
            'region': site.region,
            'nokia_pm': site.nokia_pm,
        })

    return JsonResponse({
        'sites': data,
        'total_sites': sites.count(),
        'total_by_po': total_by_po,
        'total_pending': Site.objects.filter(wcc_status='Pending').count(),
        'total_done': Site.objects.filter(wcc_status='Done').count(),
        'total_certificates': Certificate.objects.count(),
        'is_generating': request.session.get('generating', False),
    })


def api_import_logs(request):
    """Get import logs"""
    logs = ImportLog.objects.all()[:50]
    data = []
    for log in logs:
        data.append({
            'id': log.id,
            'action': log.action,
            'details': log.details,
            'sites_imported': log.sites_imported,
            'po_lines_imported': log.po_lines_imported,
            'created_at': log.created_at.strftime('%Y-%m-%d %H:%M:%S'),
            'user': log.user.username if log.user else 'System'
        })
    return JsonResponse(data, safe=False)


@csrf_exempt
def api_update_file(request):
    """Update a specific file"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        file_type_map = {
            'book2_file': 'book2',
            'purchase_orders_file': 'purchase_orders',
            'template_file': 'template'
        }

        file_type = None
        uploaded_file = None

        for key, ftype in file_type_map.items():
            if key in request.FILES:
                file_type = ftype
                uploaded_file = request.FILES[key]
                break

        if not file_type or not uploaded_file:
            return JsonResponse({'error': 'No file provided'}, status=400)

        old_file = UploadedFile.objects.filter(file_type=file_type, is_active=True).first()
        if old_file:
            try:
                old_path = os.path.join(settings.MEDIA_ROOT, old_file.file_path.name)
                if os.path.exists(old_path):
                    os.remove(old_path)
            except:
                pass
            old_file.delete()

        file_path = default_storage.save(f'uploads/{file_type}/{uploaded_file.name}', uploaded_file)
        new_file = UploadedFile.objects.create(
            file_type=file_type,
            file_name=uploaded_file.name,
            file_path=file_path,
            uploaded_by=request.user if request.user.is_authenticated else None
        )

        if file_type == 'book2':
            Site.objects.filter(source_file=old_file).delete()
            PurchaseOrder.objects.filter(source_file=old_file).delete()
            book2_abs_path = os.path.join(settings.MEDIA_ROOT, file_path)
            process_book2(book2_abs_path, new_file)
            request.session['book2_path'] = file_path
        elif file_type == 'purchase_orders':
            PurchaseOrder.objects.filter(source_file=old_file).delete()
            po_abs_path = os.path.join(settings.MEDIA_ROOT, file_path)
            process_purchase_orders(po_abs_path, new_file)
        elif file_type == 'template':
            request.session['template_path'] = file_path

        ImportLog.objects.create(
            user=request.user if request.user.is_authenticated else None,
            action='Updated File',
            details=f'Updated {uploaded_file.name}',
            sites_imported=0,
            po_lines_imported=0
        )

        return JsonResponse({
            'success': True,
            'message': f'{uploaded_file.name} updated successfully'
        })

    except Exception as e:
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)


@csrf_exempt
def api_delete_file(request, file_type):
    """Delete a specific file"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        file_obj = UploadedFile.objects.filter(file_type=file_type, is_active=True).first()
        if not file_obj:
            return JsonResponse({'error': 'File not found'}, status=404)

        try:
            file_path = os.path.join(settings.MEDIA_ROOT, file_obj.file_path.name)
            if os.path.exists(file_path):
                os.remove(file_path)
        except:
            pass

        file_obj.delete()

        if file_type == 'book2':
            Site.objects.filter(source_file=file_obj).delete()
            request.session.pop('book2_path', None)
        elif file_type == 'purchase_orders':
            PurchaseOrder.objects.filter(source_file=file_obj).delete()
        elif file_type == 'template':
            request.session.pop('template_path', None)

        ImportLog.objects.create(
            user=request.user if request.user.is_authenticated else None,
            action='Deleted File',
            details=f'Deleted {file_type} file',
            sites_imported=0,
            po_lines_imported=0
        )

        return JsonResponse({'success': True, 'message': 'File deleted'})

    except Exception as e:
        return JsonResponse({'error': str(e)}, status=500)


@csrf_exempt
def api_delete_all_files(request):
    """Delete all uploaded files"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Method not allowed'}, status=405)

    try:
        kill_excel_processes()

        for file_obj in UploadedFile.objects.all():
            try:
                if file_obj.file_path:
                    path = os.path.join(settings.MEDIA_ROOT, file_obj.file_path.name)
                    if os.path.exists(path):
                        os.remove(path)
                        print(f"✅ Deleted: {path}")
            except Exception as e:
                print(f"⚠️ Could not delete file: {e}")
            file_obj.delete()

        for cert in Certificate.objects.all():
            try:
                if cert.file_path:
                    path = os.path.join(settings.MEDIA_ROOT, cert.file_path.name)
                    if os.path.exists(path):
                        os.remove(path)
                        print(f"✅ Deleted cert: {path}")
            except Exception as e:
                print(f"⚠️ Could not delete cert file: {e}")
            try:
                if cert.pdf_file_path:
                    path = os.path.join(settings.MEDIA_ROOT, cert.pdf_file_path.name)
                    if os.path.exists(path):
                        os.remove(path)
                        print(f"✅ Deleted pdf: {path}")
            except Exception as e:
                print(f"⚠️ Could not delete pdf file: {e}")
            cert.delete()

        PurchaseOrder.objects.all().delete()
        Site.objects.all().delete()

        request.session.pop('template_path', None)
        request.session.pop('book2_path', None)
        request.session.pop('generating', None)

        ImportLog.objects.create(
            user=request.user if request.user.is_authenticated else None,
            action='Deleted All Files',
            details='All files and data deleted',
            sites_imported=0,
            po_lines_imported=0
        )

        return JsonResponse({'success': True, 'message': 'All files deleted'})

    except Exception as e:
        print(f"❌ Error in delete_all_files: {str(e)}")
        traceback.print_exc()
        return JsonResponse({'error': str(e)}, status=500)