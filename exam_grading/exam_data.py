from pathlib import Path
from typing import Optional

import qgrid 
import pandas as pd
from IPython.display import display
import numpy as np
from ipywidgets import widgets

class ExamData:
    def __init__(self, filename: str, import_from: Optional[str] = None):
        if not filename.endswith(".xlsx"):
            raise RuntimeError("File '{filename}' does not end with '.xlsx'!")

        self.path = Path(filename)
        self.path_import = Path(import_from) if import_from else None
        self.ignore_cell_edited = False
        #     filename = filename[:-1]
        # self.path_students = Path(filename+'_students.csv')
        # self.path_marks = Path(filename+'_marks.csv')
        # if not self.path_students.exists:
        #     self.path_students.parent.mkdir(parents=True, exist_ok=True)
        #     with open(self.path_students, "w") as fout:
        #         fout.write("name;total;student A; 0")
        # if not self.path_marks.exists:
        #     self.path_marks.parent.mkdir(parents=True, exist_ok=True)
        #     with open(self.path_marks, "w") as fout:
        #         fout.write("name;total;student A; 0")

        # self.df = pd.f

        events = [
            #'instance_created',
            'cell_edited',
            #'selection_changed',
            #'viewport_changed',
            'row_added',
            'row_removed',
            #'filter_dropdown_shown',
            #'filter_changed',
            #'sort_changed',
            #'text_filter_viewport_changed',
            #'json_updated'
        ]
        all_events = [
            'instance_created',
            'cell_edited',
            'selection_changed',
            'viewport_changed',
            'row_added',
            'row_removed',
            'filter_dropdown_shown',
            'filter_changed',
            'sort_changed',
            'text_filter_viewport_changed',
            'json_updated'
        ]

        self.grid_tasks = qgrid.show_grid(
            pd.DataFrame(data=dict(Task=['Task 1'],Points=[10])), 
            show_toolbar=True,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=True,
                                enableColumnReorder=False)) 
        self.grid_tasks.on(events, self.tasks_changed)

        self.grid_students = qgrid.show_grid(
            pd.DataFrame(data={'Student':['Student A'],'Mark':['4'],'Total':['5 (50%)'],'Task 1':[5]}), 
            show_toolbar=True,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=True,
                                enableColumnReorder=False,
                                column_definitions={'Total': dict(editable= False)} )) 
        self.grid_students.on(events, self.students_changed)

        self.grid_marks = qgrid.show_grid(
            pd.DataFrame(data={'Mark':['1','2','3','4','5','6'],'Min Percentage':[85,70,55,40,20,0], 'Points': [8.5,7,5.5,4,2,0]}), 
            show_toolbar=True,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=True,
                                enableColumnReorder=False )) 
        self.grid_marks.on(events, self.marks_changed)

        self.output = widgets.Output(layout={'border': '1px solid black'})

        self.load(must_exist=False)


    def _update_df(self, grid, new_df):
        if len(new_df) != len(grid.df) or list(grid.df.columns.values) != list(new_df.columns.values):
            grid.df = new_df
            return
        self.ignore_cell_edited = True
        grid_current = grid.get_changed_df()
        for col in new_df.columns.values:
            for row in range(len(new_df)):
                if grid_current.loc[row, col] != new_df.loc[row, col]:
                    grid.edit_cell(row, col, new_df.loc[row, col])
        self.ignore_cell_edited = False

    @property
    def tasks_df(self):
        return self.grid_tasks.get_changed_df()

    @tasks_df.setter
    def tasks_df(self, value):
        self._update_df(self.grid_tasks, value)

    @property
    def marks_df(self):
        return self.grid_marks.get_changed_df()

    @marks_df.setter
    def marks_df(self, value):
        value = value.astype({'Mark': 'str', 'Min Percentage': 'float', 'Points': 'float'})
        self._update_df(self.grid_marks, value)

    @property
    def students_df(self):
        return self.grid_students.get_changed_df()

    @students_df.setter
    def students_df(self, value):
        self._update_df(self.grid_students, value)


    @property
    def path_backup(self):
        return self.path.with_name(self.path.name[:-5]+'-backup.xlsx')

    def save(self):
        try:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            if self.path_backup.exists() and self.path.exists():
                self.path_backup.unlink()
            if self.path.exists():
                self.path.rename(self.path_backup)
            
            
            with pd.ExcelWriter(str(self.path)) as writer:
                self.students_df.to_excel(writer, index=False, sheet_name='Students')
                self.tasks_df.to_excel(writer, index=False, sheet_name='Tasks')
                self.marks_df.to_excel(writer, index=False, sheet_name='Marks')
        except Exception as ex:
            self.output.append_stderr(str(ex)+'\n')
        
    def load(self, must_exist=True):

        load_path = self.path
        if not load_path.exists() and self.path_import is not None:
            load_path = self.path_import
        try:
            self.students_df = pd.read_excel(str(load_path), sheet_name='Students')
            self.tasks_df = pd.read_excel(str(load_path), sheet_name='Tasks')
            self.marks_df = pd.read_excel(str(load_path), sheet_name='Marks')
        except Exception as ex:
            self.output.append_stderr(str(ex)+'\n')

    def new_taskname(self,old,new_index):
        students_df = self.students_df
        tasks_df = self.tasks_df
        new = tasks_df['Task'][new_index]
        changed = False
        iteration = 1        
        original = new
        while new in students_df:
            iteration +=1
            new = original+f'({iteration})'
            changed=True
        if old is not None:
            students_df.rename(columns={old:new}, inplace = True)
        else:
            students_df[new] = 0

        tasks_df.loc[new_index,'Task'] = new
        self.students_df = students_df
        if changed:
            self.tasks_df = tasks_df

    def recalculate_totals(self):
        students_df = self.students_df
        tasks_df = self.tasks_df
        total_points = np.sum(tasks_df['Points'])
        totals = students_df.loc[:,tasks_df['Task'].values].sum(axis=1)
        totals_perc = totals / total_points*100
        totals_perc = np.array([f'{round(p,1):.1f}' for p in totals_perc])
        students_df['Total'] = totals.astype(str)+" ("+totals_perc+'%)'
        self.students_df = students_df
        self.recalculate_marks()

    def recalculate_marks(self):
        students_df = self.students_df
        tasks_df = self.tasks_df
        marks_df = self.marks_df

        # update min points (round to half points)
        total_points = np.sum(tasks_df['Points'])
        mark_points = np.round(marks_df['Min Percentage'] / 100 * total_points * 2)/2
        marks_df['Points'] = mark_points

        # calculate student marks
        totals = students_df.loc[:,tasks_df['Task'].values].sum(axis=1)
        for row in marks_df.sort_values('Points', ascending = True).itertuples():
            mask_students = totals>=row.Points
            students_df.loc[mask_students,'Mark'] = row.Mark

        self.students_df = students_df
        self.marks_df = marks_df
        self.save()

    def tasks_changed(self, *arg, **kwarg):
        evt = arg[0]
        if self.ignore_cell_edited and evt['name'] == 'cell_edited':
            return
        if evt['name'] == 'row_added':
            self.new_taskname(None, evt['index'])
        if evt['name'] == 'cell_edited' and evt['column'] == 'Task':
            self.new_taskname(evt['old'], evt['index'])
        if evt['name'] == 'row_removed':
            students_df = self.students_df
            indices =  self.tasks_df.loc[evt['indices'], 'Task'].values
            students_df.drop(indices, axis=1, inplace=True)
            self.students_df = students_df
        self.recalculate_totals()
        # df = self.grid_students.get_changed_df()
        # print(df.columns)
        # print(self.grid_tasks.get_changed_df()['Task'].values)
        # df.columns = df.columns[:2] + self.grid_tasks.get_changed_df()['Task'].values

    def marks_changed(self, *arg, **kwarg):
        if self.ignore_cell_edited and arg[0]['name'] == 'cell_edited':
            return
        self.recalculate_marks()

    def students_changed(self, *arg, **kwarg):
        if self.ignore_cell_edited and arg[0]['name'] == 'cell_edited':
            return
        self.recalculate_totals()

    def print_event(self, *arg, **kwarg):
        evt = arg[0]
        print(f'evt: {evt}')
        if len(arg) > 2:
            print(f'arg: {arg[2:]}')
        if kwarg:
            print(f'kwarg: {kwarag}')

    def show_output(self):
        display(self.output)

    def show_students(self):
        display(self.grid_students)

    def show_tasks(self):
        display(self.grid_tasks)

    def show_marks(self):        
        display(self.grid_marks)

    def show_input(self):
        self.show_output()
        self.show_students()
        self.show_tasks()
        self.show_marks()
        
    def show_all(self):
        self.show_input()

import re