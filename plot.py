import os, sys
import matplotlib.pyplot as plt

class PlotGraph:
  def __init__(self, fname):
    self.fname = fname
    self.x = []
    self.y = []
    self.unique_data = set()

    self.read_data(self.fname)

  def read_data(self, fname):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    fpath = os.path.join(dir_path, 'input', fname)

    with open(fpath, 'r') as fh:
      for line in fh:
        self.unique_data.add(tuple(line.split()))

    for xy in self.unique_data:
      #print('{} {}'.format(xy[0], xy[1]))
      self.x.append(xy[0])
      self.y.append(xy[1])

  def show_graph(self):
    print('{} : {}'.format(len(self.x), len(self.y)))
    plt.plot(self.x, self.y, 'r')
    #plt.axis([5000, 15000, 3000, 4000])
    plt.show()

  def show_unique(self):
    for xy in self.unique_data:
      print('{} {}'.format(xy[0], xy[1]))


def test():
  pg = PlotGraph('5000_50_log.txt')
  #pg = PlotGraph('example_plot.txt')
  #pg.show_graph()
  pg.show_unique()

if __name__=="__main__":
  test()