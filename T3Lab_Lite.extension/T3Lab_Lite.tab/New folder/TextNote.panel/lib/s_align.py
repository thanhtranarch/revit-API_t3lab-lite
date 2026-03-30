# -*- coding: utf-8 -*-
"""
AlignLib

--------------------------------------------------------
Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/

--------------------------------------------------------
"""

__author__  = "Tran Tien Thanh"
__title__   = "Horizontal Right"



import sys
import os

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import *
from collections import namedtuple

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


LeaderLocation = namedtuple("LeaderLocation", ["obj","obj_leader","anchor","elbow", "end"])

class SetTextNote:
    def __init__(self, *textnotes):
        self.textnotes = textnotes
        self.cache_coordinates()
    def cache_coordinates(self):
        self.max_y = float('-inf')
        self.min_y = float('inf')
        self.max_x = float('-inf')
        self.min_x = float('inf')

        for textnote in self.textnotes:
            coord = textnote.Coord
            if coord.Y > self.max_y:
                self.max_y = coord.Y
                self.top_textnote = textnote
            if coord.Y < self.min_y:
                self.min_y = coord.Y
                self.bot_textnote = textnote
            if coord.X > self.max_x:
                self.max_x = coord.X
                self.r_textnote = textnote
            if coord.X < self.min_x:
                self.min_x = coord.X
                self.l_textnote = textnote


    def leaderset(self,location):
        # Location: 'TopLine', 'Midpoint', 'BottomLine'
        for textnote in self.textnotes:
            textnote.LeaderRightAttachment = location
            textnote.LeaderLeftAttachment = location

    def average_center(self):
        x_sum = sum(textnote.Coord.X for textnote in self.textnotes)
        return x_sum / len(self.textnotes)

    def average_middle(self):
        y_sum = sum(textnote.Coord.Y for textnote in self.textnotes)
        return y_sum / len(self.textnotes)

    def align_text(self, side, center, middle):
        v_align_options = ["left", "right", "center"]
        h_align_options = ["top", "bottom", "middle"]

        for textnote in self.textnotes:
            if side in v_align_options:
                self.vertical_align_func(side, textnote, center)
            elif side in h_align_options:
                self.horizontal_align_func(side, textnote, middle)

    def vertical_align_func(self, side, textnote, center):
        if side == "left":
            new_x = self.min_x
        elif side == "right":
            new_x = self.max_x
        elif side == "center":
            new_x = center
        textnote.Coord = XYZ(new_x, textnote.Coord.Y, textnote.Coord.Z)

    def horizontal_align_func(self, side, textnote, middle):
        if side == "top":
            new_y = self.max_y
        elif side == "bottom":
            new_y = self.min_y
        elif side == "middle":
            new_y = middle
        textnote.Coord = XYZ(textnote.Coord.X, new_y, textnote.Coord.Z)

    def distribute_hozly(self):
        number_textnote = len(self.textnotes)
        if number_textnote > 1:
            # Calculate the total width available for distribution
            total_width = self.max_x - self.min_x
            # Calculate the spacing between each text note
            spacing = total_width / (number_textnote - 1)
            # Distribute text notes horizontally
            current_x = self.min_x
            for textnote in self.textnotes:
                textnote.Coord = XYZ(current_x, textnote.Coord.Y, textnote.Coord.Z)
                current_x += spacing


    def distribute_verly(self):
        number_textnote = len(self.textnotes)
        if number_textnote > 1:
            # Calculate the total width available for distribution
            total_width = self.max_y - self.min_y
            # Calculate the spacing between each text note
            spacing = total_width / (number_textnote - 1)
            # Distribute text notes horizontally
            current_y = self.min_y
            for textnote in self.textnotes:
                textnote.Coord = XYZ(textnote.Coord.Y, current_y,textnote.Coord.Z)
                current_y += spacing

def get_leader(textnotes):
    LeaderLocation = namedtuple("LeaderLocation", ["obj", "obj_leader", "anchor", "elbow", "end"])
    leader_location = []
    for textnote in textnotes:
        leader_list = textnote.GetLeaders()
        for leader in leader_list:
            in4_leader = LeaderLocation(textnote, leader, leader.Anchor, leader.Elbow, leader.End)
            leader_location.append(in4_leader)
    return leader_location


def set_leaderdefault(leader_location, textnotes):
    obj_list=leader_location
    for textnote in textnotes:
        leader_list = textnote.GetLeaders()
        for leader in leader_list:
            for leader_exin4 in obj_list:
                if textnote == leader_exin4.obj:
                    leader.End=leader_exin4.end
                    obj_list.remove(leader_exin4)