from django.urls import path

from books import views

urlpatterns = [
    path('', views.book_create, name='book_list'),
    path('download_excel/', views.download_excel, name='download_excel'),
    # path('view/<int:pk>', views.book_view, name='book_view'),
    # path('new', views.book_create, name='book_new'),
    # path('edit/<int:pk>', views.book_update, name='book_edit'),
    # path('delete/<int:pk>', views.book_delete, name='book_delete'),
]