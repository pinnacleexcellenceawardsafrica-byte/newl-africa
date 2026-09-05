from django import forms
from .models import Site

class ImportFilesForm(forms.Form):
    book2_file = forms.FileField(
        label='Book2.xlsx',
        help_text='Site master data with PO numbers',
        widget=forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'})
    )
    purchase_orders_file = forms.FileField(
        label='Purchase Orders.xlsx',
        help_text='PO line item details',
        widget=forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'})
    )
    template_file = forms.FileField(
        label='Certificate Template.xlsx',
        help_text='Template with logos and formatting',
        widget=forms.FileInput(attrs={'class': 'form-control', 'accept': '.xlsx,.xls'})
    )


class SiteFilterForm(forms.Form):
    status = forms.ChoiceField(
        choices=[('', 'All Status')] + list(Site.STATUS_CHOICES),
        required=False,
        widget=forms.Select(attrs={'class': 'form-select'})
    )
    category = forms.ChoiceField(
        choices=[('', 'All Categories')],
        required=False,
        widget=forms.Select(attrs={'class': 'form-select'})
    )
    search = forms.CharField(
        required=False,
        widget=forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Search...'})
    )
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Get unique categories from database
        categories = Site.objects.values_list('sub_category', flat=True).distinct()
        choices = [('', 'All Categories')] + [(cat, cat) for cat in categories if cat]
        self.fields['category'].choices = choices