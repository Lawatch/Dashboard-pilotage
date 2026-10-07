"""Ouvre un .xlsm dans LibreOffice (sans fenêtre), exécute des macros VBA, enregistre en .xlsx.
Sert à tester les macros et à pré-remplir le classeur livré.

usage : python3 outils/executer_macros.py entree.xlsm sortie.xlsx Module.Macro[=argument] [...]   (sortie « - » : pas d'enregistrement)
"""
import os, subprocess, sys, tempfile, time, pathlib
import uno


def get_soffice_env():
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"
    return env


from com.sun.star.beans import PropertyValue


def P(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p


def main(src, dst, macros, timeout=600):
    profile = tempfile.mkdtemp(prefix='lo_prof_')
    port = 2002 + os.getpid() % 1000
    proc = subprocess.Popen(['soffice', '-env:UserInstallation=' + pathlib.Path(profile).as_uri(), '--headless',
                             '--invisible', '--norestore', '--nologo',
                             '--accept=socket,host=127.0.0.1,port=%d;urp;' % port],
                            env=get_soffice_env(), stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local)
        for _ in range(60):
            try:
                ctx = resolver.resolve('uno:socket,host=127.0.0.1,port=%d;urp;StarOffice.ComponentContext' % port)
                break
            except Exception:
                time.sleep(1)
        smgr = ctx.ServiceManager
        cp = smgr.createInstanceWithContext('com.sun.star.configuration.ConfigurationProvider', ctx)
        acc = cp.createInstanceWithArguments('com.sun.star.configuration.ConfigurationUpdateAccess',
                                             (P('nodepath', '/org.openoffice.Office.Calc/Filter/Import/VBA'),))
        acc.setPropertyValue('Load', True); acc.setPropertyValue('Executable', True)
        acc.commitChanges()
        desktop = smgr.createInstanceWithContext('com.sun.star.frame.Desktop', ctx)
        doc = desktop.loadComponentFromURL(pathlib.Path(src).absolute().as_uri(), '_blank', 0,
                                           (P('Hidden', True), P('MacroExecutionMode', 4)))
        sp = doc.getScriptProvider()
        for m in macros:
            t = time.time()
            try:
                nom, _, arg = m.partition('=')        # Module.Macro=argument (texte) facultatif
                s = sp.getScript('vnd.sun.star.script:VBAProject.%s?language=Basic&location=document' % nom)
                s.invoke((arg,) if arg else (), (), ())
                print('OK', m, round(time.time() - t, 1), 's', flush=True)
            except Exception as e:
                print('ERREUR', m, repr(e)[:2000], flush=True)
        if dst:
            filt = 'Calc MS Excel 2007 XML' if dst.endswith('.xlsx') else 'Calc MS Excel 2007 VBA XML'
            doc.storeToURL(pathlib.Path(dst).absolute().as_uri(), (P('FilterName', filt),))
        doc.close(True)
    finally:
        try:
            desktop.terminate()
        except Exception:
            pass
        time.sleep(1)
        proc.kill()


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2] if sys.argv[2] != '-' else None, sys.argv[3:])
