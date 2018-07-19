# -*- mode: python -*-
block_cipher = None


a = Analysis(['quick.spec'],
             pathex=['C:\\Users\\Julius\\AppData\\Local\\Programs\\Python\\Python36\\Lib\\site-packages\\scipy\\extra-dll',
                     'C:\\Users\\Julius\\Desktop\\WebTester'],
             binaries=[],
             datas=[],
             hiddenimports=['scipy._lib.messagestream', 'scipy', 'scipy.signal', 'scipy.signal.bsplines', 'scipy.special', 'scipy.special._ufuncs_cxx',
                            'scipy.linalg.cython_blas',
                            'scipy.linalg.cython_lapack',
                            'scipy.integrate',
                            'scipy.integrate.quadrature',
                            'scipy.integrate.odepack',
                            'scipy.integrate._odepack',
                            'scipy.integrate.quadpack',
                            'scipy.integrate._quadpack',
                            'scipy.integrate._ode',
                            'scipy.integrate.vode',
                            'scipy.integrate._dop', 'scipy._lib', 'scipy._build_utils','scipy.__config__',
                            'scipy.integrate.lsoda', 'scipy.cluster', 'scipy.constants','scipy.fftpack','scipy.interpolate','scipy.io','scipy.linalg','scipy.misc','scipy.ndimage','scipy.odr','scipy.optimize','scipy.setup','scipy.sparse','scipy.spatial','scipy.special','scipy.stats','scipy.version',
                            'pyodbc',
                            'sys',
                            'math',
                            'csv',
                            'os',
                            'pandas',
                            'scipy',
                            'scipy.stats',
                            'scipy.stats.stats'
                            'itertools',
                            'json',
                            'collections',
                            'pandas._libs.tslibs.timedeltas'
                            'pandas._libs.tslibs.np_datetime'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)

def get_pandas_path():
   import pandas
   pandas_path = pandas.__path__[0]
   return pandas_path


dict_tree = Tree(get_pandas_path(), prefix='pandas', excludes=["*.pyc"])
a.datas += dict_tree
a.binaries = filter(lambda x: 'pandas' not in x[0], a.binaries)

pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          exclude_binaries=True,
          name='quick',
          debug=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               name='quick')
