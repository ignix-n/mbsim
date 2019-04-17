# ===============================================================================
# Title : New Modula Consturction Process Experiment Tool(Filter)
# Date : 2019.3.12
# Author : Adriano Nam and Gabriel Nam
# Email : adriano97@gmail.com, gabriel9980@naver.com
# Description
#   - Parsing Excel File and Extract Basic Information
#   - Generate a random grouping permutation (integer partition, combination)
#   - Calcuate the group cost and property
#   - Write each group data to a new excel file
# Environment
#   - python3.x
#   - pandas, openpyxl(pip install pandas, pip install openpyxl)
# All copyright reserved to Nam's brothers(Sungyup Nam and Sunghoon Nam)
# ===============================================================================


import sys, os
import openpyxl


class GrpShortWall:
  def __init__(self, item):
    self.item = item

  def cond01(self, item):
    if 'SW1' in item:
      if 'SW2' in item:
        return True
      else:
        return False
    elif 'SW2' in item:
      if 'SW1' in item:
        return True
      else:
        return False
    else:
      return True

  def cond02(self, item):
    if 'SW3' in item:
      if 'SW4' in item:
        return True
      else:
        return False
    elif 'SW4' in item:
      if 'SW3' in item:
        return True
      else:
        return False
    else:
      return True

  def cond03(self, item):
    if len(item) >= 2:
      if 'SW5' in item:
        if 'SW3' in item and 'SW4' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond04(self, item):
    if len(item) >= 2:
      if 'SW6' in item:
        if 'SW5' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond05(self, item):
    if len(item) >= 2:
      if 'SW7' in item:
        if 'SW3' in item and 'SW4' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond06(self, item):
    if len(item) >= 2:
      if 'SW8' in item:
        if 'SW7' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond07(self, item):
    if len(item) >= 2:
      if 'SW5' in item and 'SW7' in item:
        if 'SW6' in item and 'SW9' in item and 'SW10' in item and 'SW11' in item and 'SW12' in item and 'SW13' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond08(self, item):
    if len(item) >= 2:
      if 'SW9' in item:
        if 'SW3' in item and 'SW4' in item and 'SW10' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond09(self, item):
    if len(item) >= 2:
      if 'SW10' in item:
        if 'SW3' in item and 'SW4' in item and 'SW9' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond10(self, item):
    if len(item) >= 2:
      if 'SW11' in item:
        if 'SW9' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond11(self, item):
    if len(item) >= 2:
      if 'SW12' in item:
        if 'SW1' in item and 'SW2' in item and 'SW3' in item and 'SW4' in item and 'SW13' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond12(self, item):
    if len(item) >= 2:
      if 'SW13' in item:
        if 'SW1' in item and 'SW2' in item and 'SW3' in item and 'SW4' in item and 'SW12' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond13(self, item):
    if 'SW12' in item:
      if 'SW13' in item:
        return True
      else:
        return False
    elif 'SW13' in item:
      if 'SW12' in item:
        return True
      else:
        return False
    else:
      return True

  def chk(self):
    if not self.cond01(self.item):
      return False

    if not self.cond02(self.item):
      return False

    if not self.cond03(self.item):
      return False

    if not self.cond04(self.item):
      return False

    if not self.cond05(self.item):
      return False

    if not self.cond06(self.item):
      return False

    if not self.cond07(self.item):
      return False

    if not self.cond08(self.item):
      return False

    if not self.cond09(self.item):
      return False

    if not self.cond10(self.item):
      return False

    if not self.cond11(self.item):
      return False

    if not self.cond12(self.item):
      return False

    #if not self.cond13(self.item):
    #  return False

    return True


class GrpLongWall:
  def __init__(self, item):
    self.item = item

  def cond01(self, item):
    if 'LW1' in item:
      if 'LW2' in item:
        return True
      else:
        return False
    elif 'LW2' in item:
      if 'LW1' in item:
        return True
      else:
        return False
    else:
      return True

  def cond02(self, item):
    if len(item) >= 2:
      if 'LW3' in item:
        if 'LW1' in item and 'LW2' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond03(self, item):
    if len(item) >= 2:
      if 'LW4' in item:
        if 'LW3' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond04(self, item):
    if len(item) >= 2:
      if 'LW5' in item:
        if 'LW1' in item and 'LW2' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond05(self, item):
    if len(item) >= 2:
      if 'LW3' in item and 'LW5' in item:
        if 'LW4' in item and 'LW7' in item and 'LW8' in item and 'LW9' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond06(self, item):
    if len(item) >= 2:
      if 'LW6' in item:
        if 'LW4' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond07(self, item):
    if len(item) >= 2:
      if 'LW7' in item:
        if 'LW4' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond08(self, item):
    if 'LW7' in item:
      if 'LW8' in item:
        return True
      else:
        return False
    elif 'LW8' in item:
        if 'LW7' in item:
          return True
        else:
          return False
    else:
      return True

  def cond09(self, item):
    if len(item) >= 2:
      if 'LW9' in item:
        if 'LW7' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond10(self, item):
    if len(item) >= 2:
      if 'LW8' in item:
        if 'LW1' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond11(self, item):
    if len(item) >= 2:
      if 'LW7' in item:
        if 'LW8' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond12(self, item):
    if len(item) >= 2:
      if 'LW8' in item:
        if 'LW7' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True


  def chk(self):
    if not self.cond01(self.item):
      return False

    if not self.cond02(self.item):
      return False

    if not self.cond03(self.item):
      return False

    if not self.cond04(self.item):
      return False

    if not self.cond05(self.item):
      return False

#    if not self.cond06(self.item):
#      return False

    if not self.cond07(self.item):
      return False

#    if not self.cond08(self.item):
#      return False

    if not self.cond09(self.item):
      return False

    if not self.cond10(self.item):
      return False

    if not self.cond11(self.item):
      return False

    if not self.cond12(self.item):
      return False

    return True


class GrpExtWall:
  def __init__(self, item):
    self.item = item

  def cond01(self, item):
    if 'EW1' in item :
      if 'EW2' in item and 'EW3' in item:
        return True
      else :
        return False
    return True  

#    if 'EW1' in item and 'EW2' in item and 'EW3' in item:
#      return True
#    else:
#      return False

  def chk(self):
    if not self.cond01(self.item):
      return False

    return True


class GrpBottom:
  def __init__(self, item):
    self.item = item

  def cond01(self, item):
    if 'BT1' in item:
      if 'BT2' in item:
        return True
      else:
        return False
    elif 'BT2' in item:
      if 'BT1' in item:
        return True
      else:
        return False
    else:
      return True

  def cond02(self, item):
    if len(item) >= 2:
      if 'BT3' in item:
        if 'BT1' in item and 'BT2' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond03(self, item):
    if len(item) >= 2:
      if 'BT4' in item:
        if 'BT3' in item and 'BT5' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond04(self, item):
    if len(item) >= 2:
      if 'BT5' in item:
        if 'BT3' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond05(self, item):
    if 'BT6' in item:
      if len(item) == 1:
        return True
      elif len(item) == 3:
        if 'BT7' in item and 'BT8' in item:
          return True
        else:
          return False
      else:
        return False
    else:
      return True

  def cond06(self, item):
    if 'BT7' in item:
      if len(item) == 1:
        return True
      elif len(item) == 3:
        if 'BT6' in item and 'BT8' in item:
          return True
        else:
          return False
      else:
        return False
    else:
      return True

  def cond07(self, item):
    if 'BT8' in item:
      if len(item) == 1:
        return True
      elif len(item) == 3:
        if 'BT6' in item and 'BT7' in item:
          return True
        else:
          return False
      else:
        return False
    else:
      return True

  def chk(self):
    if not self.cond01(self.item):
      return False

    if not self.cond02(self.item):
      return False

    if not self.cond03(self.item):
      return False

    if not self.cond04(self.item):
      return False

    if not self.cond05(self.item):
      return False

    if not self.cond06(self.item):
      return False

    if not self.cond07(self.item):
      return False

    return True


class GrpCeiling:
  def __init__(self, item):
    self.item = item

  def cond01(self, item):
    if 'CL1' in item:
      if 'CL2' in item:
        return True
      else:
        return False
    elif 'CL2' in item:
      if 'CL1' in item:
        return True
      else:
        return False
    else:
      return True

  def cond02(self, item):
    if len(item) >= 2:
      if 'CL3' in item:
        if 'CL1' in item and 'CL2' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond03(self, item):
    if len(item) >= 2:
      if 'CL4' in item:
        if 'CL5' in item and 'CL1' in item and 'CL2' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond04(self, item):
    if len(item) >= 2:
      if 'CL5' in item:
        if 'CL1' in item and 'CL2' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond05(self, item):
    if len(item) >= 2:
      if 'CL6' in item:
        if 'CL4' in item and 'CL5' in item and 'CL1' in item and 'CL2' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond06(self, item):
    if len(item) >= 2:
      if 'CL7' in item:
        if 'CL4' in item and 'CL6' in item and 'CL9' in item and 'CL10' in item and 'CL3' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond07(self, item):
    if len(item) >= 2:
      if 'CL8' in item:
        if 'CL1' in item and 'CL2' in item and 'CL3' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond08(self, item):
    if len(item) >= 2:
      if 'CL9' in item:
        if 'CL1' in item and 'CL2' in item and 'CL3' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond09(self, item):
    if len(item) >= 2:
      if 'CL10' in item:
        if 'CL9' in item:
          return True
        else:
          return False
      else:
        return True
    else:
      return True

  def cond10(self, item):
    if len(item) >= 2:
      if 'CL4' in item:
        if 'CL5' in item:
          if 'CL6' in item:
            return True
          else:
            return False
        else:
          return False
      elif 'CL5' in item:
        if 'CL4' in item:
          if 'CL6' in item:
            return True
          else:
            return False
        else:
          return False
      elif 'CL6' in item:
        if 'CL4' in item:
          if 'CL5' in item:
            return True
          else:
            return False
        else:
          return False
      else:
        return True
    else:
      return True

  def chk(self):
    if not self.cond01(self.item):
      return False

    if not self.cond02(self.item):
      return False

    if not self.cond03(self.item):
      return False

    if not self.cond04(self.item):
      return False

    if not self.cond05(self.item):
      return False

    if not self.cond06(self.item):
      return False

    #if not self.cond07(self.item):
    #  return False

    if not self.cond08(self.item):
      return False

    if not self.cond09(self.item):
      return False

    if not self.cond10(self.item):
      return False

    return True


class GrpPipeRack:
  def __init__(self, item):
    self.item = item

  def chk(self):
    return True


class GrpToiletWall:
  def __init__(self, item):
    self.item = item

  def chk(self):
    if len(self.item) == 1 or len(self.item) == 11:
      return True
    else:
      if len(self.item) == 8:
        if not 'TW8' in self.item and not 'TW10' in self.item and not 'TW11' in self.item:
          return True
        else:
          return False
      else:
        return False
