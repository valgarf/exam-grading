from pathlib import Path
from typing import Optional

import qgrid 
import pandas as pd
from IPython.display import display,HTML
import numpy as np
from ipywidgets import widgets, Button, HBox, VBox
from matplotlib import pyplot as plt
from matplotlib import gridspec
from xlrd import XLRDError

from ipywidgets import Button, HBox, VBox

class ExamData:
    def __init__(self, filename: str, export_to:str, import_from: Optional[str] = None, grade_postfixes = {'+':-0.3, '-':0.3}):
        if not filename.endswith(".xlsx"):
            raise RuntimeError("File '{filename}' does not end with '.xlsx'!")
        if not export_to.endswith(".xlsx"):
            raise RuntimeError("File '{export_to}' does not end with '.xlsx'!")

        self.path = Path(filename)
        self.path_export = Path(export_to)
        self.path_import = Path(import_from) if import_from else None
        self.ignore_cell_edited = False
        self.grade_postfixes = grade_postfixes
        self._disable_update=False
        self._undo = []
        self._redo = []
        self._students_fixed_cols=4
        
        events = [
            'cell_edited',
            'row_added',
            'row_removed',
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
        self.grid_tasks.on(['selection_changed','row_removed'],self.select_last_row)

        self.grid_students = qgrid.show_grid(
            pd.DataFrame(data={'Student':['Student A'],'Grade':['4'],'Grade Correction': [0],'Total':['5 (50%)'],'Task 1':[5]}), 
            show_toolbar=True,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=True,
                                enableColumnReorder=False,
                                column_definitions={'Total': dict(editable= False)} )) 
        self.grid_students.on(events, self.students_changed)
        self.grid_students.on(['selection_changed','row_removed'],self.select_last_row)

        self.grid_grades = qgrid.show_grid(
            pd.DataFrame(data={'Grade':['1','2','3','4','5','6'],'Min Percentage':[85,70,55,40,20,0], 'Min Points': [8.5,7,5.5,4,2,0], 'Min Points': [10,8,6.5,5,3.5,1.5]}), 
            show_toolbar=True,
            grid_options = dict(maxVisibleRows = 30,
                                minVisibleRows = 2,
                                sortable=False,
                                filterable=False,
                                autoEdit=True,
                                enableColumnReorder=False )) 
        self.grid_grades.on(events, self.grades_changed)
        self.grid_grades.on(['selection_changed','row_removed'],self.select_last_row)
        

        self.grid_score = qgrid.show_grid(
            pd.DataFrame(data={'Grade':[],'Amount':[], 'Students': []}), 
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
            pd.DataFrame(data={'Points':[], 'Grade': [], 'Amount':[], 'Students': []}), 
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
        self.button_update = Button(description='Updating enabled', button_style='success', tooltip="Toggles the automatic recalculation when editing cells in the table. These auto updates might be annoying when trying to edit multiple data fields.")
        self.button_postfix = widgets.ToggleButton(
                value=False,
                description='Drop Postfix',
                disabled=False,
                button_style='info', # 'success', 'info', 'warning', 'danger' or ''
                tooltip="Wether or not the grade overview and plot will aggregate grades with different postfixes into one grade. E.g. the grades '2-','2','2+' will all be shown as '2'.",
                icon='times' # (FontAwesome names without the `fa-` prefix)
            )
        self.passed_perc = widgets.FloatText(
                    value=40,
                    description='Passed:',
                    disabled=False
                )

        self.button_export.on_click(self.export)
        self.button_undo.on_click(self.undo)
        self.button_redo.on_click(self.redo)
        self.button_update.on_click(self.switch_disable_update)
        self.button_postfix.observe(self.switch_aggregate_postfix)
        self.passed_perc.observe(self.passed_percentage_changed)
        self.button_bar = HBox([self.button_undo, self.button_redo, self.button_export, self.button_update,self.button_postfix, self.passed_perc, widgets.Label(value='%')])

        self.load(must_exist=False)
        

    def export(self, btn):
        self.save(export=True)

    def _set_current_state(self):
        self._current_state = {
            'tasks': self.tasks_df,
            'grades': self.grades_df,
            'students': self.students_df
        }

    def _restore_state(self, state):
        self.tasks_df = state['tasks'] 
        self.grades_df = state['grades'] 
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

    def switch_disable_update(self, *arg):
        self._disable_update =  not self._disable_update
        if self._disable_update:
            self.button_update.button_style='warning'
            self.button_update.description='Updating disabled'
        else:
            self.button_update.button_style='success'
            self.button_update.description='Updating enabled'
            self.recalculate_totals()

    def switch_aggregate_postfix(self, *arg):
        if not arg[0]['name'] == 'value':
            return
        if self.button_postfix.value:
            self.button_postfix.icon = 'check'
        else:
            self.button_postfix.icon = 'times'
        self.recalculate_output()

    def passed_percentage_changed(self, *arg):
        if not arg[0]['name'] == 'value':
            return
        self.recalculate_output()


    def _update_df(self, grid, new_df):
        grid.df = new_df
        return

        # grid_current = grid.get_changed_df()
        # if list(new_df.index.values) != list(grid_current.index.values) or list(new_df.columns.values) != list(grid_current.columns.values):
        #     grid.df = new_df
        #     # print('complete update')
        #     return
        # self.ignore_cell_edited = True
        # for col in new_df.columns.values:
        #     for row in range(len(new_df)):
        #         if grid_current.loc[row, col] != new_df.loc[row, col]:
        #             grid.edit_cell(row, col, new_df.loc[row, col])
        #             # print(f'update: {row}, {col}')
        # self.ignore_cell_edited = False

    @property
    def tasks_df(self):
        return self.grid_tasks.get_changed_df()

    @tasks_df.setter
    def tasks_df(self, value):
        self._update_df(self.grid_tasks, value)

    @property
    def grades_df(self):
        return self.grid_grades.get_changed_df()

    @grades_df.setter
    def grades_df(self, value):
        value = value.astype({'Grade': 'str', 'Min Percentage': 'float', 'Min Points': 'float', 'Max Points': 'float'})
        self._update_df(self.grid_grades, value)

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

            additional = pd.DataFrame(data=dict(postfix=[self.button_postfix.value], passed=[self.passed_perc.value]))            
            with pd.ExcelWriter(str(path)) as writer:
                self.students_df.to_excel(writer, index=False, sheet_name='Students')
                self.tasks_df.to_excel(writer, index=False, sheet_name='Tasks')
                self.grades_df.to_excel(writer, index=False, sheet_name='Grades')
                additional.to_excel(writer, index=False, sheet_name='Additional')
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
            students_df = pd.read_excel(str(load_path), sheet_name='Students')
            tasks_df = pd.read_excel(str(load_path), sheet_name='Tasks')
            try:
                grades_df = pd.read_excel(str(load_path), sheet_name='Grades')
                additional = pd.read_excel(str(load_path), sheet_name='Additional')
                self.button_postfix.value = bool(additional['postfix'][0])
                self.passed_perc.value = additional['passed'][0]
            except XLRDError:
                grades_df = pd.read_excel(str(load_path), sheet_name='Marks')
                students_df = students_df.rename(columns={'Mark':'Grade'})
                grades_df = grades_df.rename(columns={'Mark':'Grade'})
                students_df.insert(2,'Adjustment',0)
            self.students_df = students_df
            self.grades_df = grades_df
            self.tasks_df = tasks_df
            self.recalculate_totals()    
        except Exception as ex:
            self.output_text.append_stderr(str(ex)+'\n')
            raise
        

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
        if self._disable_update:
            self._set_current_state()
            self.save()
            return
        students_df = self.students_df
        tasks_df = self.tasks_df
        total_points = np.sum(tasks_df['Points'])
        totals = students_df.loc[:,tasks_df['Task'].values].sum(axis=1)
        totals_perc = totals / total_points*100
        totals_perc = np.array([f'{round(p,1):.1f}' for p in totals_perc])
        students_df['Total'] = totals.astype(str)+" ("+totals_perc+'%)'
        self.students_df = students_df
        self.recalculate_grades()

    def recalculate_grades(self):
        if self._disable_update:
            self._set_current_state()
            self.save()
            return
        students_df = self.students_df
        tasks_df = self.tasks_df
        grades_df = self.grades_df

        # update min points (round to half points)
        total_points = np.sum(tasks_df['Points'])
        grade_points = np.round(grades_df['Min Percentage'] / 100 * total_points * 2)/2
        grades_df['Min Points'] = grade_points
        grades_df['Max Points'] = np.array([total_points,] + list(grade_points - 0.5)[:-1])

        # calculate student grades
        totals = students_df.loc[:,tasks_df['Task'].values].sum(axis=1)
        grade_list= [] 
        for row in grades_df.sort_values('Min Points', ascending = True).rename(columns={'Min Points': 'Points'}).itertuples():
            grade_list.append(row.Grade)
            mask_students = totals>=row.Points
            students_df.loc[mask_students,'Grade'] = row.Grade

        def adjust_grade(student):
            if student.Adjustment != 0:
                idx = grade_list.index(student.Grade)
                idx += student.Adjustment
                if idx < 0:
                    idx = 0
                if idx >= len(grade_list):
                    idx = -1
                student.Grade = grade_list[idx]
            return student
        students_df = students_df.apply(adjust_grade, axis=1)

        self.students_df = students_df
        self.grades_df = grades_df
        self.recalculate_output()

    def _grade_to_float(self, grades):
        result = []
        for g in grades:
            corr = 0
            for pf, pf_corr in self.grade_postfixes.items():
                if g.endswith(pf):
                    g = g[:-len(pf)]
                    corr += pf_corr
            result.append(float(g) + corr)
        return np.array(result)

    def _strip_postfix(self, grades):
        result = []
        for g in grades:
            corr = 0
            for pf, pf_corr in self.grade_postfixes.items():
                if g.endswith(pf):
                    g = g[:-len(pf)]
                    corr += pf_corr
            result.append(g)
        return np.array(result)

    def recalculate_output(self):
        self._set_current_state()
        self.save()

        # prepare students
        students_df = self.students_df.copy()
        students_df['Points'] = students_df.loc[:,self.tasks_df['Task'].values].sum(axis=1)
        students_df['Amount'] = 1
        students_df['Students'] = students_df['Student']
      
        # aggregate points / score dataframes
        def agg_names(seq):
            return ", ".join(seq.values)

        self.grid_points.df = (students_df
                                .groupby('Points')
                                .agg({'Grade': 'first', 'Amount': 'sum', 'Students': agg_names}))
        score_base_df = students_df
        if self.button_postfix.value:
            score_base_df = students_df.copy()
            score_base_df['Grade'] = self._strip_postfix(score_base_df['Grade'])
        self.grid_score.df = (score_base_df
                                .groupby('Grade')
                                .agg({'Amount': 'sum', 'Students': agg_names}))
                                
        with self.output_plot_score:
            self.output_plot_score.clear_output(True)
            self._plot_score()

        with self.output_info:
            self.output_info.clear_output(True)
            self._print_output()

    def _print_output(self):
        # prepare students 
        total_points = np.sum(self.tasks_df['Points'])   
        students_df = self.students_df.copy()
        students_df['Points'] = students_df.loc[:,self.tasks_df['Task'].values].sum(axis=1)
        students_df['Percentage'] = students_df['Points']*100 / total_points
        passed = students_df[students_df['Percentage']>=self.passed_perc.value]
        failed = students_df[students_df['Percentage']<self.passed_perc.value]

        # print information
        print('Information')
        print('-----------\n')
        df = self.grid_score.df
        mean = np.sum(self._grade_to_float(df.index.values) * df['Amount'])/np.sum(df['Amount'])
        print(f'Ø {mean}')
        print(f'average percentage: {np.mean(students_df["Percentage"]):.1f}%')
        print(f'median percentage: {np.median(students_df["Percentage"]):.1f}%')
        print(f'#students: {len(self.students_df)}')
        print(f'#passed: {len(passed)} ({len(passed)*100/len(self.students_df):.1f}%), #failed: {len(failed)} ({len(failed)*100/len(self.students_df):.1f}%)')
        print(f'passing percentage: {self.passed_perc.value}%')
        if not self.passed_perc.value in self.grades_df['Min Percentage'].values:
            print(f'WARNING: passing percentage does not coincide with a grade!')

    def _plot_score(self):
        # prepare students table with additional precalculated information        
        total_points = np.sum(self.tasks_df['Points'])      
        students_df = self.students_df.copy()
        students_df['Points'] = students_df.loc[:,self.tasks_df['Task'].values].sum(axis=1)
        students_df.loc[students_df['Points'] > total_points,'Points'] = total_points
        students_df['Amount'] = 1
        students_df['Students'] = students_df['Student']
      
        # create plotting dataframe for points        
        points = np.arange(0, total_points+0.1,step=0.5)
        plot_points = pd.DataFrame(index=points, data=dict(Amount=0))
        agg_points = (students_df
                        .groupby('Points')
                        .agg({'Grade': 'first', 'Amount': 'sum'}))
        plot_points['Amount'] = agg_points['Amount']

        # create plotting dataframe for score
        grade_list = self.grades_df['Grade'].values
        if self.button_postfix.value:
            grade_list = np.unique(self._strip_postfix(grade_list))
        plot_score = pd.DataFrame(index=grade_list, data=dict(Amount=0))
        plot_score['Amount'] = self.grid_score.df['Amount']

        # plot
        fig = plt.figure(figsize=(17, 5)) 
        gs = gridspec.GridSpec(1, 2, width_ratios=[3, 4]) 
        ax1 = plt.subplot(gs[0])
        ax2 = plt.subplot(gs[1])
        plot_score.plot.bar(y='Amount', color='#28d9ed', ax=ax1)
        plot_points.plot.bar(y='Amount', color='#28d9ed', ax=ax2)
        xticks_pos = (self.grades_df['Min Points'].values + self.grades_df['Max Points'].values)
        xticks_labels = self.grades_df['Grade'].values
        plt.xticks(list(xticks_pos),list(xticks_labels))
        grades_no_postfix = self.grades_df.copy()
        grades_no_postfix['Grade'] = self._strip_postfix(grades_no_postfix['Grade'])
        xlines_large = grades_no_postfix.groupby('Grade').agg({'Min Points': 'min'})['Min Points'].values[:-1]
        xlines = self.grades_df['Min Points'].values[:-1]
        xlines = np.array([xl for xl in xlines if xl not in xlines_large])
        plt.vlines((xlines-0.25)*2, 0, np.max(plot_points['Amount']*0.45), linestyles='dotted')
        plt.vlines((xlines_large-0.25)*2, 0, np.max(plot_points['Amount']*0.8), linestyles='dotted')
        plt.show()
        
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
            cols = self.students_df.columns.values[self._students_fixed_cols:]
            indices = [cols[i] for i in evt["indices"]]
            students_df.drop(indices, axis=1, inplace=True)
            self.students_df = students_df
            self.tasks_df = self.tasks_df.reset_index(drop=True)
        self.recalculate_totals()


    def grades_changed(self, *arg, **kwarg):
        if self.ignore_cell_edited and arg[0]['name'] == 'cell_edited':
            return
        self._add_to_history()
        self.recalculate_grades()

    def students_changed(self, *arg, **kwarg):
        if self.ignore_cell_edited and arg[0]['name'] == 'cell_edited':
            return
        self._add_to_history()
        self.recalculate_totals()

    def select_last_row(self, evt, grid):
        if not evt['name'] == 'selection_changed' or (not evt['new'] and evt['source'] == 'gui'):
            df = grid.get_changed_df()
            grid.change_selection(rows=[len(df)-1])
        
    def print_event(self, *arg, **kwarg):
        with self.output_text:
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

    def show_grades(self):        
        display(self.grid_grades)

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
        self.show_grades()
        
    def show_all(self):
        self.init()
        self.show_input()
        self.show_output()

    def show_output(self):
        display(self.output_info)
        display(self.output_plot_score)
        display(self.grid_score)
        display(self.grid_points)
        
    def show_printable(self):
        with pd.option_context("display.max_rows", None, "display.max_columns", None):
            display(self.students_df.set_index('Student'))
            display(self.tasks_df.set_index('Task'))
            display(self.grades_df.set_index('Grade'))
            self._print_output()
            self._plot_score()
            display(self.grid_score.df)
            display(self.grid_points.df)

import re