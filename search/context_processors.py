from search.models import Codemaster

def common(request):
    context = {
        'skills':Codemaster.objects.all
    }
    return context