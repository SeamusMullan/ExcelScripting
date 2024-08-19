from typing import Union, Tuple, List, Optional, Any

import customtkinter as ctk

from customtkinter.windows.widgets.core_rendering import CTkCanvas
from customtkinter.windows.widgets.theme import ThemeManager
from customtkinter.windows.widgets.core_rendering import DrawEngine
from customtkinter.windows.widgets.core_widget_classes import CTkBaseClass

import functions as fns

class DictionaryFrame(CTkBaseClass):
    
    # This class represents a frame with 2 columns of n text boxes
    # It represents a dictionary / key+value structure
    
    
    def __init__(self,
                 master: Any,
                 width: int = 200,
                 height: int = 200,
                 corner_radius: Optional[Union[int, str]] = None,
                 border_width: Optional[Union[int, str]] = None,

                 bg_color: Union[str, Tuple[str, str]] = "transparent",
                 fg_color: Optional[Union[str, Tuple[str, str]]] = None,
                 border_color: Optional[Union[str, Tuple[str, str]]] = None,

                 background_corner_colors: Union[Tuple[Union[str, Tuple[str, str]]], None] = None,
                 overwrite_preferred_drawing_method: Union[str, None] = None,
                 num_rows: int = 10,
                 num_cols: int = 2,
                 **kwargs):

        # transfer basic functionality (_bg_color, size, __appearance_mode, scaling) to CTkBaseClass
        super().__init__(master=master, bg_color=bg_color, width=width, height=height, **kwargs)


        self.num_rows = num_rows
        self.num_cols = num_cols
        fns.log(f"{self.num_rows} Rows and {self.num_cols} Columns", 'message')
        self.text_boxes = []


        # color
        self._border_color = ThemeManager.theme["CTkFrame"]["border_color"] if border_color is None else self._check_color_type(border_color)

        # determine fg_color of frame
        if fg_color is None:
            if isinstance(self.master, DictionaryFrame):
                if self.master._fg_color == ThemeManager.theme["CTkFrame"]["fg_color"]:
                    self._fg_color = ThemeManager.theme["CTkFrame"]["top_fg_color"]
                else:
                    self._fg_color = ThemeManager.theme["CTkFrame"]["fg_color"]
            else:
                self._fg_color = ThemeManager.theme["CTkFrame"]["fg_color"]
        else:
            self._fg_color = self._check_color_type(fg_color, transparency=True)

        self._background_corner_colors = background_corner_colors  # rendering options for DrawEngine

        # shape
        self._corner_radius = ThemeManager.theme["CTkFrame"]["corner_radius"] if corner_radius is None else corner_radius
        self._border_width = ThemeManager.theme["CTkFrame"]["border_width"] if border_width is None else border_width

        self._canvas = CTkCanvas(master=self,
                                 highlightthickness=0,
                                 width=self._apply_widget_scaling(self._current_width),
                                 height=self._apply_widget_scaling(self._current_height))
        self._canvas.place(x=0, y=0, relwidth=1, relheight=1)
        self._canvas.configure(bg=self._apply_appearance_mode(self._bg_color))
        self._draw_engine = DrawEngine(self._canvas)
        self._overwrite_preferred_drawing_method = overwrite_preferred_drawing_method

        self._draw(no_color_updates=True)
        self._create_text_boxes()

    def winfo_children(self) -> List[any]:
        """
        winfo_children of DictionaryFrame without self.canvas widget,
        because it's not a child but part of the DictionaryFrame itself
        """

        child_widgets = super().winfo_children()
        try:
            child_widgets.remove(self._canvas)
            return child_widgets
        except ValueError:
            return child_widgets

    def _set_scaling(self, *args, **kwargs):
        super()._set_scaling(*args, **kwargs)

        self._canvas.configure(width=self._apply_widget_scaling(self._desired_width),
                               height=self._apply_widget_scaling(self._desired_height))
        self._draw()

    def _set_dimensions(self, width=None, height=None):
        super()._set_dimensions(width, height)

        self._canvas.configure(width=self._apply_widget_scaling(self._desired_width),
                               height=self._apply_widget_scaling(self._desired_height))
        self._draw()


    def set_num_rows(self, num):
        self.num_rows = num
        
    def set_num_cols(self, num):
        self.num_cols = num

# Draw text boxes
    def _create_text_boxes(self):
        col_names = ["Input", "Output"]
    
        # Create column labels
        for col in range(self.num_cols):
            label = ctk.CTkLabel(self, text=col_names[col] if col < len(col_names) else f"Column {col+1}")
            label.grid(row=0, column=col, padx=5, pady=5)
            
        for row in range(self.num_rows):
            row_boxes = []
            for col in range(self.num_cols):
                text_box = ctk.CTkEntry(self, width=self._current_width // self.num_cols, height=30)
                text_box.grid(row=row+1, column=col, padx=5, pady=5)
                row_boxes.append(text_box)
            self.text_boxes.append(row_boxes)
        # self._canvas.tag_lower("inner_parts")  # maybe unnecessary, I don't know ???
        # self._canvas.tag_lower("border_parts")

    def get_data(self) -> dict:
        data = {}
        for row in self.text_boxes:
            key = row[0].get()
            value = row[1].get()
            if key:
                data[key] = value
        return data

    def set_data(self, data: dict):
        for i, (key, value) in enumerate(data.items()):
            if i < self.num_rows:
                self.text_boxes[i][0].insert(0, key)
                self.text_boxes[i][1].insert(0, value)


    def _draw(self, no_color_updates=False):
        super()._draw(no_color_updates)

        if not self._canvas.winfo_exists():
            return

        if self._background_corner_colors is not None:
            self._draw_engine.draw_background_corners(self._apply_widget_scaling(self._current_width),
                                                      self._apply_widget_scaling(self._current_height))
            self._canvas.itemconfig("background_corner_top_left", fill=self._apply_appearance_mode(self._background_corner_colors[0]))
            self._canvas.itemconfig("background_corner_top_right", fill=self._apply_appearance_mode(self._background_corner_colors[1]))
            self._canvas.itemconfig("background_corner_bottom_right", fill=self._apply_appearance_mode(self._background_corner_colors[2]))
            self._canvas.itemconfig("background_corner_bottom_left", fill=self._apply_appearance_mode(self._background_corner_colors[3]))
        else:
            self._canvas.delete("background_parts")

        requires_recoloring = self._draw_engine.draw_rounded_rect_with_border(self._apply_widget_scaling(self._current_width),
                                                                              self._apply_widget_scaling(self._current_height),
                                                                              self._apply_widget_scaling(self._corner_radius),
                                                                              self._apply_widget_scaling(self._border_width),
                                                                              overwrite_preferred_drawing_method=self._overwrite_preferred_drawing_method)

        if no_color_updates is False or requires_recoloring:
            if self._fg_color == "transparent":
                self._canvas.itemconfig("inner_parts",
                                        fill=self._apply_appearance_mode(self._bg_color),
                                        outline=self._apply_appearance_mode(self._bg_color))
            else:
                self._canvas.itemconfig("inner_parts",
                                        fill=self._apply_appearance_mode(self._fg_color),
                                        outline=self._apply_appearance_mode(self._fg_color))

            self._canvas.itemconfig("border_parts",
                                    fill=self._apply_appearance_mode(self._border_color),
                                    outline=self._apply_appearance_mode(self._border_color))
            self._canvas.configure(bg=self._apply_appearance_mode(self._bg_color))
            
    
    def configure(self, require_redraw=False, **kwargs):
        if "fg_color" in kwargs:
            self._fg_color = self._check_color_type(kwargs.pop("fg_color"), transparency=True)
            require_redraw = True

            # check if CTk widgets are children of the frame and change their bg_color to new frame fg_color
            for child in self.winfo_children():
                if isinstance(child, CTkBaseClass):
                    child.configure(bg_color=self._fg_color)

        if "bg_color" in kwargs:
            # pass bg_color change to children if fg_color is "transparent"
            if self._fg_color == "transparent":
                for child in self.winfo_children():
                    if isinstance(child, CTkBaseClass):
                        child.configure(bg_color=self._fg_color)

        if "border_color" in kwargs:
            self._border_color = self._check_color_type(kwargs.pop("border_color"))
            require_redraw = True

        if "background_corner_colors" in kwargs:
            self._background_corner_colors = kwargs.pop("background_corner_colors")
            require_redraw = True

        if "corner_radius" in kwargs:
            self._corner_radius = kwargs.pop("corner_radius")
            require_redraw = True

        if "border_width" in kwargs:
            self._border_width = kwargs.pop("border_width")
            require_redraw = True

        super().configure(require_redraw=require_redraw, **kwargs)

    def cget(self, attribute_name: str) -> any:
        if attribute_name == "corner_radius":
            return self._corner_radius
        elif attribute_name == "border_width":
            return self._border_width

        elif attribute_name == "fg_color":
            return self._fg_color
        elif attribute_name == "border_color":
            return self._border_color
        elif attribute_name == "background_corner_colors":
            return self._background_corner_colors

        else:
            return super().cget(attribute_name)

    def bind(self, sequence=None, command=None, add=True):
        """ called on the tkinter.Canvas """
        if not (add == "+" or add is True):
            raise ValueError("'add' argument can only be '+' or True to preserve internal callbacks")
        self._canvas.bind(sequence, command, add=True)

    def unbind(self, sequence=None, funcid=None):
        """ called on the tkinter.Canvas """
        if funcid is not None:
            raise ValueError("'funcid' argument can only be None, because there is a bug in" +
                             " tkinter and its not clear whether the internal callbacks will be unbinded or not")
        self._canvas.unbind(sequence, None)

    