jupyter contrib nbextension install --user
jupyter nbextension install --py widgetsnbextension --user
jupyter nbextension enable --py --user widgetsnbextension
jupyter nbextension install --py qgrid --user
jupyter nbextension enable --py --user qgrid
jupyter nbextension install --py ipysheet --user
jupyter nbextension enable --py --user ipysheet
jupyter nbextension enable init_cell/main
python -m nbopen.install_win