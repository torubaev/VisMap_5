import pyvista as pv

mesh = pv.Sphere()
plotter = pv.Plotter()
plotter.add_mesh(mesh)
plotter.show()