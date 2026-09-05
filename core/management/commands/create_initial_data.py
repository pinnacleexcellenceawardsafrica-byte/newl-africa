from django.core.management.base import BaseCommand
from content.models import WebsiteContent

class Command(BaseCommand):
    help = 'Create initial website content'

    def handle(self, *args, **options):
        # Create initial content
        initial_content = [
            # Hero Section
            ('site_name', 'Mwasamwanda Wellness', 'hero'),
            ('site_title', 'Mwasamwanda - Mental Health Services', 'hero'),
            ('hero_tagline', 'Empowerment, Enhanced Productivity, Happy Life', 'hero'),
            ('hero_title', 'Enhancing Normative Development and Mental Health Across the Lifespan', 'hero'),
            ('hero_description', 'Mwasamwanda Well-being Services facilitates normative development and mental health through professional, evidence-based psychological interventions for individuals and groups.', 'hero'),
            
            # Services Section
            ('services_title', 'Professional Psychological Services', 'service'),
            ('services_subtitle', 'Evidence-based interventions designed to enhance normative development and mental health across the lifespan', 'service'),
            ('service1_title', 'Consultancy & Advisory', 'service'),
            ('service1_price', 'KSh 5,000', 'service'),
            ('service1_description', 'Professional psychological consultancy including policy formulation, program planning, research design, and psychological assessment.', 'service'),
            ('service2_title', 'Counselling & Psychotherapy', 'service'),
            ('service2_price', 'KSh 3,500/session', 'service'),
            ('service2_description', 'Professional therapy services across the lifespan to help individuals deal with issues limiting normative development and mental health.', 'service'),
            ('service3_title', 'Training Programs', 'service'),
            ('service3_price', 'KSh 8,000 - 15,000', 'service'),
            ('service3_description', 'Comprehensive training programs to build capacity in organizations, institutions, and communities on normative development and mental health.', 'service'),
            ('service4_title', 'Mind Transformation Coaching', 'service'),
            ('service4_price', 'KSh 6,000/session', 'service'),
            ('service4_description', 'Specialized coaching services focusing on personal development, mindset transformation, and enhanced productivity.', 'service'),
            
            # Contact Section
            ('contact_title', 'Contact Mwasamwanda Well-being Services', 'contact'),
            ('contact_subtitle', 'Get in touch with us to discuss your mental health needs or schedule a consultation.', 'contact'),
            ('phone_number', '0758283613', 'contact'),
            ('email_address', 'mwasawellservices@gmail.com', 'contact'),
            ('address_line1', 'Southern House, Murang\'a Road', 'contact'),
            ('address_line2', 'Off Mot Avenue, Fourth Floor', 'contact'),
            ('address_line3', 'Nairobi, KENYA', 'contact'),
            ('hours_weekdays', 'Mon-Fri: 8:00 AM - 6:00 PM', 'contact'),
            ('hours_weekend', 'Sat: 9:00 AM - 2:00 PM', 'contact'),
        ]

        created_count = 0
        updated_count = 0

        for key, value, content_type in initial_content:
            obj, created = WebsiteContent.objects.get_or_create(
                key=key,
                defaults={
                    'value': value,
                    'content_type': content_type,
                    'description': f'Content for {key}'
                }
            )
            
            if created:
                created_count += 1
                self.stdout.write(self.style.SUCCESS(f'Created: {key}'))
            else:
                # Update existing content
                obj.value = value
                obj.content_type = content_type
                obj.description = f'Content for {key}'
                obj.save()
                updated_count += 1
                self.stdout.write(self.style.WARNING(f'Updated: {key}'))

        self.stdout.write(
            self.style.SUCCESS(
                f'Successfully created {created_count} and updated {updated_count} content items!'
            )
        )