from django.urls import path
from search.views import top
from search.views import (EngineerListView,
                          SearchSkillEngineerListView,
                          SearchNameEngineerListView,
                          SearchSomethingProjectListView,
                          ProjectListView,
                          SearchProjectByEngineerName,
                          EngineerDetailView,
                          ProjectDetailView,
                          basic_upload,
                          error
                          )
from search import views

#TOP画面
urlpatterns = [
    path('', top, name='top'),
#案件情報登録
    path('project_fileupload', basic_upload, name='project_fileupload'),
#案件、技術者検索処理
    path('engineer',EngineerListView.as_view(),name='engineer-list'),
    path('engineer_skill',SearchSkillEngineerListView.as_view(),name='skill-engineer-list'),
    path('engineer_name',SearchNameEngineerListView.as_view(),name='name-engineer-list'),
    path('project',ProjectListView.as_view(),name = "project-list"),
    path('some',SearchSomethingProjectListView.as_view(),name='some-project-list'),
    path('skillforproject',SearchProjectByEngineerName.as_view(),name='engineername-project-list'),  
#技術者詳細画面
    path('engineer/engineer_detail/<int:pk>', EngineerDetailView.as_view(), name='engineer_detail'),
#案件詳細画面
    path('project/project_detail/<int:pk>', ProjectDetailView.as_view(), name='project_detail'),
#技術者登録
    path('engineer_file', views.engineer_upload, name='engineer'),
#エラー画面
    path('error', error, name='error')
]