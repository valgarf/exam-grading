from pathlib import Path
from typing import Optional

import qgrid 
import pandas as pd
from IPython.display import display,HTML
import numpy as np
from ipywidgets import widgets, Button, HBox, VBox
from matplotlib import pyplot as plt
from matplotlib import gridspec

from ipywidgets import Button, HBox, VBox

class ExamData:
    def __init__(self, filename: str, export_to:str, import_from: Optional[str] = None):
        if not filename.endswith(".xlsx"):
            raise RuntimeError("File '{filename}' does not end with '.xlsx'!")
        if not export_to.endswith(".xlsx"):
            raise RuntimeError("File '{export_to}' does not end with '.xlsx'!")

        self.path = Path(filename)
        self.path_export = Path(export_to)
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
            pd.DataFrame(data={'Mark':['1','2','3','4','5','6'],'Min Percentage':[85,70,55,40,20,0], 'Min Points': [8.5,7,5.5,4,2,0], 'Min Points': [10,8,6.5,5,3.5,1.5]}), 
            show_toolbar=True,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=True,
                                enableColumnReorder=False )) 
        self.grid_marks.on(events, self.marks_changed)

        self.grid_score = qgrid.show_grid(
            pd.DataFrame(data={'Mark':[],'Amount':[], 'Students': []}), 
            show_toolbar=False,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=False,
                                editable=False,
                                enableColumnReorder=False,
                                )) 

        self.grid_points = qgrid.show_grid(
            pd.DataFrame(data={'Points':[], 'Mark': [], 'Amount':[], 'Students': []}), 
            show_toolbar=False,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=False,
                                editable=False,
                                enableColumnReorder=False,
                                )) 

        self.output_text = widgets.Output(layout={'border': '1px solid black'})
        self.output_plot_score = widgets.Output()
        self.output_info = widgets.Output()

        self.button_undo = Button(description='Undo')
        self.button_redo = Button(description='Redo')
        self.button_export = Button(description='Export')
        self.button_export.on_click(self.export)
        self.button_undo.on_click(self.undo)
        self.button_redo.on_click(self.redo)
        self.button_bar = HBox([self.button_undo, self.button_redo, self.button_export])

        self._undo = []
        self._redo = []

        self.load(must_exist=False)
        

    def export(self, btn):
        self.save(export=True)

    def _set_current_state(self):
        self._current_state = {
            'tasks': self.tasks_df,
            'marks': self.marks_df,
            'students': self.students_df
        }

    def _restore_state(self, state):
        self.tasks_df = state['tasks'] 
        self.marks_df = state['marks'] 
        self.students_df = state['students'] 
        self.recalculate_totals()

    def _add_to_history(self):
        self._undo.append(self._current_state)
        self._redo = []
        self._set_current_state()

    def undo(self, *arg):
        if not self._undo:
            return
        self._set_current_state()
        self._redo.append(self._current_state)
        self._restore_state(self._undo.pop())
        self._set_current_state()

    def redo(self, *arg):
        if not self._redo:
            return
        self._set_current_state()
        self._undo.append(self._current_state)
        self._restore_state(self._redo.pop())
        self._set_current_state()


    def _update_df(self, grid, new_df):
        grid_current = grid.get_changed_df()
        if len(new_df) != len(grid_current) or list(new_df.columns.values) != list(grid_current.columns.values):
            grid.df = new_df
            # print('complete update')
            return
        self.ignore_cell_edited = True
        for col in new_df.columns.values:
            for row in range(len(new_df)):
                if grid_current.loc[row, col] != new_df.loc[row, col]:
                    grid.edit_cell(row, col, new_df.loc[row, col])
                    # print(f'update: {row}, {col}')
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
        value = value.astype({'Mark': 'str', 'Min Percentage': 'float', 'Min Points': 'float', 'Max Points': 'float'})
        self._update_df(self.grid_marks, value)

    @property
    def students_df(self):
        return self.grid_students.get_changed_df()

    @students_df.setter
    def students_df(self, value):
        self._update_df(self.grid_students, value)

    def save(self, export=False):
        path = self.path_export if export else self.path
        try:
            path_backup = path.with_name(path.name[:-5]+'-backup.xlsx')
            path.parent.mkdir(parents=True, exist_ok=True)
            if path_backup.exists() and path.exists():
                path_backup.unlink()
            if path.exists():
                path.rename(path_backup)            
            
            with pd.ExcelWriter(str(path)) as writer:
                self.students_df.to_excel(writer, index=False, sheet_name='Students')
                self.tasks_df.to_excel(writer, index=False, sheet_name='Tasks')
                self.marks_df.to_excel(writer, index=False, sheet_name='Marks')
                if export:
                    self.grid_points.df.to_excel(writer, index=True, sheet_name='Points')
                    self.grid_score.df.to_excel(writer, index=True, sheet_name='Score')
        except Exception as ex:
            self.output_text.append_stderr(str(ex)+'\n')
        
    def load(self, must_exist=True):
        load_path = self.path
        if not load_path.exists() and self.path_import is not None:
            load_path = self.path_import
        try:
            self.students_df = pd.read_excel(str(load_path), sheet_name='Students')
            self.tasks_df = pd.read_excel(str(load_path), sheet_name='Tasks')
            self.marks_df = pd.read_excel(str(load_path), sheet_name='Marks')
            self.recalculate_totals()    
        except Exception as ex:
            self.output_text.append_stderr(str(ex)+'\n')
        

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
            self.grid_tasks.df = tasks_df

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
        marks_df['Min Points'] = mark_points
        marks_df['Max Points'] = np.array([total_points,] + list(mark_points - 0.5)[:-1])

        # calculate student marks
        totals = students_df.loc[:,tasks_df['Task'].values].sum(axis=1)
        for row in marks_df.sort_values('Min Points', ascending = True).rename(columns={'Min Points': 'Points'}).itertuples():
            mask_students = totals>=row.Points
            students_df.loc[mask_students,'Mark'] = row.Mark

        self.students_df = students_df
        self.marks_df = marks_df
        self._set_current_state()
        self.save()
        self.recalculate_output()

    def recalculate_output(self):
        students_df = self.students_df.copy()
        students_df['Points'] = students_df.loc[:,self.tasks_df['Task'].values].sum(axis=1)
        students_df['Amount'] = 1
        students_df['Students'] = students_df['Student']

        def agg_names(seq):
            return ", ".join(seq.values)

        self.grid_points.df = (students_df
                                .groupby('Points')
                                .agg({'Mark': 'first', 'Amount': 'sum', 'Students': agg_names}))
        self.grid_score.df = (students_df
                                .groupby('Mark')
                                .agg({'Amount': 'sum', 'Students': agg_names}))

        # plot_points= self.grid_points.df.reset_index()[['Points','Amount']]
        total_points = np.sum(self.tasks_df['Points'])
        points = np.arange(0, total_points+0.1,step=0.5)
        plot_points = pd.DataFrame(index=points, data=dict(Amount=0))
        truncated_points = students_df.copy()
        truncated_points.loc[truncated_points['Points'] > total_points,'Points'] = total_points
        truncated_points = (truncated_points
                                .groupby('Points')
                                .agg({'Mark': 'first', 'Amount': 'sum', 'Students': agg_names}))
        plot_points['Amount'] = truncated_points['Amount']

        plot_score = pd.DataFrame(index=self.marks_df['Mark'].values, data=dict(Amount=0))
        plot_score['Amount'] = self.grid_score.df['Amount']
        
        with self.output_plot_score:
            self.output_plot_score.clear_output(True)
            fig = plt.figure(figsize=(17, 5)) 
            gs = gridspec.GridSpec(1, 2, width_ratios=[3, 4]) 
            ax1 = plt.subplot(gs[0])
            ax2 = plt.subplot(gs[1])
            plot_score.plot.bar(y='Amount', color='#28d9ed', ax=ax1)
            plot_points.plot.bar(y='Amount', color='#28d9ed', ax=ax2)
            xticks_pos = (self.marks_df['Min Points'].values + self.marks_df['Max Points'].values)
            xticks_labels = self.marks_df['Mark'].values
            plt.xticks(list(xticks_pos),list(xticks_labels))
            xlines = list((self.marks_df['Min Points'].values-0.25)*2)[:-1]
            plt.vlines(xlines, 0, np.max(plot_points['Amount']*0.75), linestyles='dotted')
            plt.show()

        with self.output_info:
            self.output_info.clear_output(True)
            print('Information')
            print('-----------\n')
            df = self.grid_score.df
            mean = np.sum(df.index.astype(int) * df['Amount'])/np.sum(df['Amount'])
            print(f'Ø {mean}')
            print(f'#students: {len(self.students_df)}')

    def tasks_changed(self, *arg, **kwarg):
        evt = arg[0]
        if self.ignore_cell_edited and evt['name'] == 'cell_edited':
            return
        self._add_to_history()
        if evt['name'] == 'row_added':
            self.new_taskname(None, evt['index'])
        if evt['name'] == 'cell_edited' and evt['column'] == 'Task':
            self.new_taskname(evt['old'], evt['index'])
        if evt['name'] == 'row_removed':
            students_df = self.students_df
            cols = self.students_df.columns.values[3:]
            indices = [cols[i] for i in evt["indices"]]
            students_df.drop(indices, axis=1, inplace=True)
            self.students_df = students_df
        self.recalculate_totals()


    def marks_changed(self, *arg, **kwarg):
        if self.ignore_cell_edited and arg[0]['name'] == 'cell_edited':
            return
        self._add_to_history()
        self.recalculate_marks()

    def students_changed(self, *arg, **kwarg):
        if self.ignore_cell_edited and arg[0]['name'] == 'cell_edited':
            return
        self._add_to_history()
        self.recalculate_totals()

    def print_event(self, *arg, **kwarg):
        evt = arg[0]
        print(f'evt: {evt}')
        if len(arg) > 2:
            print(f'arg: {arg[2:]}')
        if kwarg:
            print(f'kwarg: {kwarag}')

    def show_output_text(self):
        display(self.output_text)

    def show_students(self):
        display(self.grid_students)

    def show_tasks(self):
        display(self.grid_tasks)

    def show_marks(self):        
        display(self.grid_marks)

    def init(self):
        display(HTML("""
        <style>
        .slick-header-column {
            background-color: rgb(255, 214, 90) !important;
        }
        .slick-resizable-handle {
            border-left-color: rgb(255, 214, 90) !important;
            border-right-color: rgb(255, 214, 90) !important;
            background-color: rgb(141, 119, 50) !important;
        }
        </style>
        """))

    def show_buttons(self):
        display(self.button_bar)

    def show_input(self):
        self.show_output_text()
        self.show_buttons()
        self.show_students()
        self.show_tasks()
        self.show_marks()
        
    def show_all(self):
        self.init()
        self.show_input()
        self.show_output()

    def show_output(self):
        display(self.output_info)
        display(self.output_plot_score)
        display(self.grid_score)
        display(self.grid_points)
        
        

import re