from .models import SiteSettings, Category, Nominee, Partner

def site_settings(request):
    settings = SiteSettings.objects.first()
    if not settings:
        settings = SiteSettings.objects.create()
    return {'site_settings': settings}

def global_stats(request):
    return {
        'total_categories': Category.objects.filter(is_active=True).count(),
        'total_nominees': Nominee.objects.filter(status='active').count(),
        'total_partners': Partner.objects.filter(is_active=True).count(),
    }