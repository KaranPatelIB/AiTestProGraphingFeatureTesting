import graphviz
import fitz
import textwrap
from PIL import Image, ImageFont, ImageDraw


# -------- Create Graph --------

def make_graph():
    graph = graphviz.Digraph("ex")

    # invisible nodes to add margin
    graph.node('header', "", shape='none')
    graph.node('footer', "", shape='none')

    graph.node('head', "http://testaws.dgsms.ca/", shape='note', URL="#page=1")
    graph.node('A1', '#page=1', shape='cds', color='chartreuse', href="http://testaws.dgsms.ca/")
    graph.node('A2', 'Trial Registration', shape='cds', color='chartreuse')

    with graph.subgraph(name="Cluster_B1") as subGraph:
        # subGraph.attr()
        subGraph.node('B1', 'username', shape='underline')
        subGraph.node('B2', 'userpass', shape='underline')
        subGraph.node('B3', 'submit Login', shape='cds', color='chartreuse')
        subGraph.edges([('B1', 'B2'), ('B2', 'B3')])

    graph.node('C1', 'http://testaws.dgsms.ca/LoginAction?action=login', shape='note')

    graph.edges([('head', 'A1'), ('head', 'A2'), ('head', 'B1'), ('B3', 'C1')])

    # make edges invisible
    graph.edge('header', 'head', style="invis", minlen="1")
    graph.edge('C1', 'footer', style="invis", minlen="1")

    graph.render('graphExample', view=True)

make_graph()