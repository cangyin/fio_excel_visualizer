from email.policy import default
import re
import atexit
import win32com as win32
from time import time
from types import SimpleNamespace
from win32com.client import constants

# The Excel Application COM object.
app = win32.client.gencache.EnsureDispatch('Excel.Application')


class ExcelUtils():

    def __init__(self, workbook):
        # frequently used COM objects
        self.workbook = workbook
        self.sheets = workbook.Worksheets

        atexit.register(lambda : self.set_interactive(True))

        # Messy regex below. for parsing VBA Function's and Sub's.
        # to test regex expressions: https://regex101.com/
        # TODO: move to lexer.
        accessmodifier = r'(?:Public|Protected|Friend|Private|Protected Friend)'
        proceduremodifiers = r'(?:Overloads|Overrides|Overridable|NotOverridable|MustOverride|MustOverride Overrides|NotOverridable Overrides)'
        comma_delimited_list = r'[^,\s]+(\s*,\s*[^,\s]+)*'

        _vba_function_pattern = rf'(<[^>]+>\s+)?({accessmodifier}\s+)?({proceduremodifiers}\s+)?(Shared\s+)?(Shadows\s+)?'
        _vba_function_pattern += rf'Function\s+(?P<proc_name>[0-9a-zA-Z_]+)\s*(\(Of (?P<typeparamlist>[^\)]+)\))?(\((?P<parameterlist>[^\)]*)\))?\s*(As\s+(?P<returntype>[^\s]+)\s+)?((Implements|Handles)\s+{comma_delimited_list}\s+)?(?P<proc_body>.+)\s+End\s+Function'
        self._vba_function_pattern = re.compile(_vba_function_pattern, flags=re.IGNORECASE|re.DOTALL)
        
        _vba_sub_pattern = rf'(<[^>]+>\s+)?(Partial\s+)?({accessmodifier}\s+)?({proceduremodifiers}\s+)?(Shared\s+)?(Shadows\s+)?'
        _vba_sub_pattern += rf'Sub\s+(?P<proc_name>[0-9a-zA-Z_]+)\s*(\(Of (?P<typeparamlist>[^\)]+)\))?(\((?P<parameterlist>[^\)]*)\))?\s*((Implements|Handles)\s+{comma_delimited_list}\s+)?(?P<proc_body>.+)\s*End\s+Sub'
        self._vba_sub_pattern = re.compile(_vba_sub_pattern, flags=re.IGNORECASE|re.DOTALL)

        # print(_vba_function_pattern, '\n', _vba_sub_pattern)
        self._vba_pattern_field_names = ['proc_name', 'parameterlist', 'returntype', 'proc_body']


    def get_sheet(self, name):
        try:
            sheet = self.sheets(name)
        except:
            sheet = self.sheets.Add()
            sheet.Name = name
        return sheet


    def set_interactive(self, b):
        app.Visible = b
        app.ScreenUpdating = b
        app.DisplayAlerts = b


    def range(self, rbox):
        '''
        rbox:
            The range box, a pair of coordinates. As
            [ (row_start, col_start), (row_end, col_end) ]
        '''
        s, e = rbox
        sheet = self.workbook.ActiveSheet
        return sheet.Range(sheet.Cells(*s), sheet.Cells(*e))


    def column_end(self, col :int) -> int:
        '''
        return the 'row number' of the last cell that contains data in the given column.
        '''
        sheet = self.workbook.ActiveSheet
        return sheet.Cells(sheet.Rows.Count, col).End(constants.xlUp).Row


    def row_end(self, row :int) -> int:
        '''
        return the 'column number' of the last cell that contains data in the given row.
        '''
        sheet = self.workbook.ActiveSheet
        return sheet.Cells(row, sheet.Columns.Count).End(constants.xlToLeft).Column


    def fill_cells(self, start_row, start_col, arr):
        '''
        arr : 2-dimentional array, list of lists
        '''
        if len(arr) == 0 or len(arr[0]) == 0:
            return
        rng = self.range([ (start_row, start_col), (start_row + len(arr) - 1, start_col + len(arr[0]) - 1) ])
        rng.Value = arr


    def fill_column(self, start_row, start_col, l :list):
        '''
        l : 1-dimentional list
        '''
        self.fill_cells(start_row, start_col, [[v] for v in l])


    def fill_row(self, start_row, start_col, l :list):
        '''
        l : 1-dimentional list
        '''
        self.fill_cells(start_row, start_col, [l])


    def set_formula(self, rbox, formula :str):
        '''
        set R1C1-style formula in the given range.
        '''
        rng = self.range(rbox)
        cell = rng.Cells(1, 1)
        cell.FormulaR1C1 = formula
        cell.AutoFill(rng, constants.xlFillDefault)


    def _parse_vba_proc(self, code :str):
        '''
        Parse VBA code and return proc name, parameters, return type
        (of function), and proc body.

        If the code is not inside "Sub ... End Sub" or "Function ...
        End Function", it will be wrapped by "Sub ... End Sub".
        '''
        # TODO: move to lexer, see http://www.dabeaz.com/ply/
        code = code.strip()
        if len(code) == 0 or not isinstance(code, str):
            return None

        tokens = code.lower().split()
        if not ( 'sub' in tokens or 'function' in tokens ):
            code = f'Sub sub_{int(time()*1000)}\n' + code + '\nEnd Sub'
            tokens = ['sub']

        # empty parentheses are annoying
        _code = code
        _m_empty_paren = re.search(r'\(\s*\)', code)
        if _m_empty_paren:
            s, e = _m_empty_paren.span(0)
            _code = _code[:s] + '*'*(e-s) + _code[e:]

        m = None
        if 'sub' in tokens:
            m = self._vba_sub_pattern.match(_code)
        elif 'function' in tokens:
            m = self._vba_function_pattern.match(_code)

        result = None
        if m:
            result = SimpleNamespace(**{ n: None for n in self._vba_pattern_field_names })
            result.code = code
            result.__dict__.update({ n: code[m.start(n):m.end(n)] for n in self._vba_pattern_field_names if n in m.groupdict() })
        return result


    def inject_vba_proc(self, code :str) -> str:
        '''
        Parse the VBA code, add it to VBProject of the workbook, and returns the
        procedure name.
        '''
        parsed = self._parse_vba_proc(code)
        self.vbmodule.CodeModule.AddFromString(parsed.code)
        return parsed.proc_name


    def run_vba(self, proc :str, is_proc_name :bool=False):
        '''
        run VBA code, or a VBA procedure.

        proc:
            the VBA code to run, or name of the VBA procedure to run.
        '''
        if is_proc_name:
            proc_name = proc
        else:
            proc_name = self.inject_vba_proc(proc)
        return app.Run(self.vbmodule.Name + '.' + proc_name)


    def delete_vba_proc(self, proc_name :str):
        start = self.vbmodule.CodeModule.ProcStartLine(proc_name, constants.vbext_pk_Proc) # 0 = vbext_pk_Proc
        count = self.vbmodule.CodeModule.ProcCountLines(proc_name, 0) # 0 = vbext_pk_Proc
        self.vbmodule.CodeModule.DeleteLines(start, count)


    @property
    def vbmodule(self):
        if not hasattr(self, '_vbmodule'):
            self._vbmodule_name = 'python_controlled'
            try: # try exsiting
                self._vbmodule = self.workbook.VBProject.VBComponents(self._vbmodule_name)
                self._vbmodule.CodeModule.DeleteLines(1, self._vbmodule.CodeModule.CountOfLines)
            except: # or add new
                self._vbmodule = self.workbook.VBProject.VBComponents.Add(1) # vbext_ct_StdModule
                self._vbmodule.Name = self._vbmodule_name
        return self._vbmodule


    @vbmodule.setter
    def vbmodule(self, vbcomponent):
        if vbcomponent.Type != 1:
            raise ValueError('Component type must be vbext_ct_StdModule (1).')
        self._vbmodule = vbcomponent
        self._vbmodule_name = vbcomponent.Name


    @staticmethod
    def no_screen_updating(func):
        '''
        decorator to prevent screen update.
        '''
        def wrapper(*args, **kwargs):
            old_val = app.ScreenUpdating
            app.ScreenUpdating = False
            result = func(*args, **kwargs)
            app.ScreenUpdating = old_val
            return result
        return wrapper


class ExcelXYChartUtils():
    '''
    Excel chart manipulation helper utility class, mainly focused on X-Y charts.
    '''

    def __init__(
            self,
            sheet,
            title=None,
            category_range=None,
            value_range=None,
            chart_type=constants.xlXYScatterLines,
            series_name=None
        ):
        '''
        Create a new empty chart, add a series if both category_range and
        value_range are given, or you can call add_series() later.

        sheet:
            the sheet to add the chart to.
        title:
            the chart title.
        category_range, value_range, chart_type, series_name:
            see add_series().
        '''
        self.util = ExcelUtils(sheet.Parent)
        # -1 for default chart style, see https://docs.microsoft.com/en-us/office/vba/api/excel.shapes.addchart2
        self.chart = sheet.Shapes.AddChart2(-1, constants.xlXYScatterLines).Chart
        self.chart.ChartArea.ClearContents()
        self.set_title(title)
        if category_range and value_range:
            self.add_series(category_range, value_range, chart_type, series_name)

    def delete(self):
        self.chart.Delete()

    def set_position(self, left, top):
        self.chart.Parent.Left = left
        self.chart.Parent.Top = top

    @ExcelUtils.no_screen_updating
    def set_title(self, title):
        title = str(title)
        if title:
            self.chart.HasTitle = True
            self.chart.ChartTitle.Text = title
        else:
            self.chart.HasTitle = False
        return self

    def _set_axis_title(self, axis, title :str, text_orientation=constants.xlHorizontal):
        title = str(title)
        if title:
            axis.HasTitle = True
            axis.AxisTitle.Text = title
            axis.AxisTitle.Orientation = text_orientation
        else:
            axis.HasTitle = False

    def set_x_title(self, title):
        # share the same title between primary and secondary x-axis.
        axis = self.axis(constants.xlCategory, constants.xlPrimary)
        self._set_axis_title(axis, title)
        return self

    def axis(self, axis_type=constants.xlCategory, axis_group=constants.xlPrimary):
        if not any([ s.AxisGroup == axis_group for s in self.chart.FullSeriesCollection() ]):
            # raise exception here, not to dive into win32com
            raise Exception(f'No series belongs to axis group {axis_group}.')
        return self.chart.Axes(axis_type, axis_group)

    def set_y_title(self, primary=None, secondary=None, text_orientation=constants.xlHorizontal):
        
        for title, axis_group in ((primary, constants.xlPrimary), (secondary, constants.xlSecondary)):
            if title:
                axis = self.axis(constants.xlValue, axis_group)
                self._set_axis_title(axis, title, text_orientation)
        return self

    def set_x_tick_format(self, format=None):
        """
        for example: '0.00_', "#,##0_);[红色](#,##0)"
        """
        if format and isinstance(format, str):
            axis = self.axis(constants.xlCategory, constants.xlPrimary)
            axis.TickLabels.NumberFormatLocal = format
        return self

    def set_y_tick_format(self, primary=None, secondary=None):
        for format, axis_group in ((primary, constants.xlPrimary), (secondary, constants.xlSecondary)):
            if format and isinstance(format, str):
                try:
                    axis = self.axis(constants.xlValue, axis_group)
                    axis.TickLabels.NumberFormatLocal = format
                except:
                    pass
        return self

    def set_legend_visible(self, visible=True, position=constants.xlLegendPositionTop):
        self.chart.HasLegend = visible
        if visible:
            self.chart.Legend.Position = position
        return self

    def add_series(
            self,
            category_range,
            value_range,
            chart_type=constants.xlXYScatterLines,
            axis_group=constants.xlPrimary,
            name :str=None
        ):
        '''
        category_range, value_range:
            a range object, or start and end corrdinate of a range.
        '''
        if isinstance(category_range, list):
            category_range = self.util.range(category_range)
        if isinstance(value_range, list):
            value_range = self.util.range(value_range)

        if axis_group != 1 and self.chart.FullSeriesCollection().Count == 0:
            axis_group = 1

        series = self.chart.SeriesCollection().NewSeries()
        series.XValues = category_range
        series.Values = value_range
        series.ChartType = chart_type
        series.AxisGroup = axis_group
        if name and isinstance(name, str):
            series.Name = name
        return self

    def add_series2(self, triple_range :list, **kwargs):
        '''
        triple_range:
            a list of 3 integers, (row-start, column1-start, column2-start).
            category range is (row-start, column1-start) to (row-end-of-column1, column1-start),
            value range is (row-start, column2-start) to (row-end-of-column1, column2-start).
        kwargs:
            see add_series().
        '''
        if len(triple_range) != 3:
            raise ValueError('triple_range must be a list of 3 integers.')
        row1, col1, col2 = triple_range
        row2 = self.util.column_end(col1)
        self.add_series([(row1, col1), (row2, col1)], [(row1, col2), (row2, col2)], **kwargs)


    @classmethod
    @ExcelUtils.no_screen_updating
    def create_with_declaration(cls, sheet, decl :dict):
        '''
        # Chart Declaration
        {
            "title": "chart title",

            # use 'series' or 'series_range'. if both are given, all series will be added.
            "series":  [
                # see add_series2()
                [ (row-start, column1-start, column2-start), axis-group,  "series_name1" ],
                [ (1, 2, 3), 1,  "series_name2" ],
            ],
            "series_range":  [
                # see add_series()
                [ (category_range1, value_range1), axis-group,  "series_name1" ],
                [ (category_range2, value_range2), 2,  "series_name2" ],
            ],
            "x_title": "x_title",
            "y_titles": [ "y_title_primary", "y_title_secondary" ],
            "tick_format": "tick_format",
            
            "position": [x, y],

            # VBA code to run in context of 'With ActiveChart ... End With'.
            "with_chart": """.FullSeriesSlection(1).Format.Line.Weight = 5"""
        }
        '''
        if not decl:
            return None

        self = cls(sheet, decl.get("title", None))

        # weak check
        if "series" in decl and all([ len(s) == 3 for s in decl["series"] ]):
            for s in decl["series"]:
                if len(s[0]) != 3:
                    raise Exception("Invalid series declaration.")
                self.add_series2(s[0], axis_group=s[1], name=s[2])

        if "series_range" in decl and all([ len(s) == 3 for s in decl["series_range"] ]):
            for s in decl["series_range"]:
                if len(s[0]) != 2:
                    raise Exception("Invalid series range declaration.")
                self.add_series(*s[0], axis_group=s[1], name=s[2])

        self.set_x_title(decl.get("x_title", None))

        primary, secondary = decl.get("y_titles", ['', ''])
        self.set_y_title(primary, None, constants.xlVertical if len(primary) > 10 else constants.xlHorizontal)
        self.set_y_title(None, secondary, constants.xlVertical if len(secondary) > 10 else constants.xlHorizontal)

        tick_format = decl.get("tick_format", "#,##0_);[红色](#,##0)")
        self.set_x_tick_format(tick_format)
        self.set_y_tick_format(tick_format, tick_format)

        if "position" in decl:
            x, y = decl["position"]
            self.set_position(x, y)

        if "with_chart" in decl:
            self.chart.Select()
            vba = "With ActiveChart\n" + decl["with_chart"] + "\nEnd With"
            self.util.run_vba(vba)

        self.set_legend_visible()

        return self


if __name__ == "__main__":
    import os
    version = app.Version # '16.0'
    cmd = f'reg add HKCU\SOFTWARE\Microsoft\Office\{version}\Excel\Security /v AccessVBOM /d 1 /t REG_DWORD /reg:64 /f'

    print("Enabling access to Visual Basic Object Model, with command: \n", cmd)
    os.system(cmd)


def vba_test(workbook):
    from pprint import pprint
    '''
    # Definition of VBA Function:
    [ <attributelist> ] [ accessmodifier ] [ proceduremodifiers ] [ Shared ] [ Shadows ] 
    Function name [ (Of typeparamlist) ] [ (parameterlist) ] [ As returntype ] [ Implements implementslist | Handles eventlist ]
        [ statements ]
        [ Exit Function ]
        [ statements ]
    End Function

    # Definition of VBA Sub:
    [ <attributelist> ] [ Partial ] [ accessmodifier ] [ proceduremodifiers ] [ Shared ] [ Shadows ] 
    Sub name [ (Of typeparamlist) ] [ (parameterlist) ] [ Implements implementslist | Handles eventlist ]
        [ statements ]
        [ Exit Sub ]
        [ statements ]
    End Sub
    '''

    sub1 = '''
    Sub computeArea  (ByVal length As Double, ByVal width As Double)
        ' Declare local variable.
        Dim area As Double
        If length = 0 Or width = 0 Then
            ' If either argument = 0 then exit Sub immediately.
            Exit Sub
        End If
        ' Calculate area of rectangle.
        area = length * width
        ' Print area to Immediate window.
        Debug.WriteLine(area)
    End Sub
    '''
    func1='''
    Public Function calcSum(ByVal ParamArray args() As Double) As Double
        calcSum = 0
        If args.Length <= 0 Then Exit Function
        For i As Integer = 0 To UBound(args, 1)
            calcSum += args(i)
        Next i
    End Function
    '''
    util = ExcelUtils(workbook)

    # print(util._vba_sub_pattern.pattern, '\n')
    # print(util._vba_function_pattern.pattern)

    pprint(
        util._parse_vba_proc('call_something').__dict__
    )
    print(
        util._vba_sub_pattern.match('''Sub sub_123()\ncall_something\nEnd Sub''').groupdict()
    )

    parsed = util._parse_vba_proc(sub1)
    pprint(parsed.__dict__, indent=4, sort_dicts=False, compact=False)

    parsed = util._parse_vba_proc(func1)
    pprint(parsed.__dict__, indent=4, sort_dicts=False, compact=False)

# vba_test(app.ActiveWorkbook)


def vba_test2():

    workbook = app.ActiveWorkbook
    util = ExcelUtils(workbook)

    code = '''sub TestMacro()
        msgbox "Testing 1 2 3"
        end sub
    '''
    util.inject_vba_proc('''
    Sub Test()
        Dim a As Integer
        a = 1
        Debug.Print a
        msgbox "Hello World! a=" & a

    End Sub
    ''')

    name = util.inject_vba_proc('''
    Sub computeArea  (ByVal length As Double, ByVal width As Double)
        ' Declare local variable.
        Dim area As Double
        If length = 0 Or width = 0 Then
            ' If either argument = 0 then exit Sub immediately.
            Exit Sub
        End If
        ' Calculate area of rectangle.
        area = length * width
        ' Print area to Immediate window.
        Debug.WriteLine(area)
    End Sub'''
    )
    name = util.inject_vba_proc('''
    Sub computeAre2 (ByVal length As Double, ByVal width As Double)
        ' Declare local variable.
        Dim area As Double
        If length = 0 Or width = 0 Then
            ' If either argument = 0 then exit Sub immediately.
            Exit Sub
        End If
        ' Calculate area of rectangle.
        area = length * width
        ' Print area to Immediate window.
        Debug.WriteLine(area)
    End Sub'''
    )
    print(name, 'injected')
    util.run_vba('Test')
    util.delete_proc('computeArea')

# vba_test2()


def fill_range_test(workbook):
    sheets = workbook.Worksheets
    util = ExcelUtils(workbook)

    big_table = [
        [None, 2, 3, 4, 5, 6, 7, None, 9, 10], [11, "Yyds", 13, 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "cos", 37, 38, 39, 40], [41, 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, "Yyds", 55, 56, 57, 58, 59, 60], [61, "Yyds", 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, 74, 75, 76, 77, 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, "Yyds", 14, 15, 16, 17, 18, "Yyds", 20], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, 36, 37, 38, 39, 40], [41, 42, 43, "Yyds", 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, "Yyds", 59, 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, 74, 75, 76, 77, 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20], [21, 22, "Yyds", 24, 25, 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, 36, 37, 38, 39, 40], [41, 42, 43, 44, 45, "Yyds", 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, 59, 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, 74, 75, 76, 77, 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, "Yyds", 13, 14, 15, "Yyds", 17, 18, 19, 20], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, 36, 37, 38, 39, 40], [41, 42, 43, 44, 45, 46, 47, "Yyds", 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, 59, 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, 74, 75, 76, 77, 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [31, 32, 33, "Yyds", 35, 36, 37, 38, 39, 40], [41, 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, "Yyds", 59, 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, "Yyds", 74, 75, 76, 77, 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, 36, 37, 38, 39, 40], [41, 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, "Yyds", 54, 55, 56, 57, 58, 59, 60], [61, 62, 63, 64, 65, 66, "Yyds", 68, 69, 70], [71, 72, 73, 74, 75, 76, 77, 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [31, 32, "Yyds", 34, 35, 36, "Yyds", 38, 39, 40], [41, 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, "Yyds", 56, 57, 58, 59, 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [71, 72, 73, 74, 75, 76, 77, 78, 79, 80], [1, 2, 3, 4, 5, 6, 7, 8, 9, 10], [11, 12, 13, 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, 36, 37, 38, 39, 40], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [21, 22, 23, 24, 25, 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, 36, 37, 38, 39, 40], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80], [11, 12,"yyds", 14, 15, 16, 17, 18, 19, 20], [21, 22, 23, 24, "Yyds", 26, 27, 28, 29, 30], [31, 32, 33, 34, 35, "Yyds", 37, 38, 39, 40], ["Yyds", 42, 43, 44, 45, 46, 47, 48, 49, 50], [51, 52, 53, 54, 55, 56, 57, 58, "Yyds", 60], [61, 62, 63, 64, 65, 66, 67, 68, 69, 70], [71, 72, 73, "Yyds", 75, 76, "Yyds", 78, 79, 80],
    ]
    print(f"filling cells from array: big_table({len(big_table)}, {len(big_table[0])})")

    iters = 600
    start = time()

    for i in range(iters):
        util.fill_cells(1, 1, big_table)
    end = time()

    elapsed = end - start
    print(f'elapsed {elapsed :.2} s, {elapsed / iters :.4} s/iteration.', )

# fill_range_test(app.ActiveWorkbook)


def chart_test(workbook):
    
    sheets = workbook.Worksheets
    sheet = sheets('tmp2')
    
    chart_util = ExcelXYChartUtils(sheet, 'yyds_title')
    chart_util.add_series( [(4, 1), (78, 1)], [(4, 2), (78, 2)], name='yyds_series', axis_group=1 )
    chart_util.add_series( [(4, 1), (78, 1)], [(4, 6), (78, 6)], name='yyds_series', axis_group=2 )

    # legend
    chart_util.set_legend_visible(True, constants.xlLegendPositionCorner)

    # x, y title
    chart_util.set_x_title('x_title')
    chart_util.set_y_title(primary='y_title_primary', secondary='y_title_secondary')

    # x, y tick format
    chart_util.set_x_tick_format()
    chart_util.set_y_tick_format()

# chart_test(app.ActiveWorkbook)


def chart_test2(workbook):
    
    sheets = workbook.Worksheets
    sheet = sheets('tmp2')
    
    # chart_util = ExcelXYChartUtils(sheet, 'yyds_title')
    chart_util = ExcelXYChartUtils.create_with_declaration(sheet, {
        "title": "Read Performance",
        "series":  [
            [ (4, 1, 2), 1,  "Bandwidth" ],
            [ (4, 1, 6), 2,  "IOPS" ],
        ],
        "x_title": "msec",
        "y_titles": [ "KB/s", "IO/s" ],
    })
    # app.ScreenUpdating = True

# chart_test2(app.ActiveWorkbook)
