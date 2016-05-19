from PySide.QtCore import *
from PySide.QtGui import *
import time
import sys

# import gtk
# from PyQt4 import QtCore, QtGui
# import ctypes
# import pyodbc
# import os
# from os import path
# import getpass
# import struct
# import socket
# import netifaces
# from netifaces import interfaces, ifaddresses, AF_INET, AF_INET6, AF_LINK
# import csv
# import dxfgrabber
# from objbrowser import browse

###
app = QApplication(sys.argv)

try:
    due = QTime.currentTime()
    message = "Alert!"

    if len(sys.argv) < 2:
        raise ValueError

    hours, minutes = sys.argv[1].split(":")
    due = QTime(int(hours), int(minutes))

    if not due.isValid():
        raise ValueError

    if len(sys.argv) > 2:
        message = " ".join(sys.argv[2:])

except ValueError:
    message = "Usage: py_odbc_GUI.py HH:MM"  # 24hr clock

while QTime.currentTime() < due:
    time.sleep(10)

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

'''
try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)
'''


def QButton():
    try:

        ex = Ui_Form()
        ex.show()
        app.exec_()
    except KeyboardInterrupt:
        print "\n\Program aborted by user...\n"
        sys.exit()


class Main_UI_Form(QtGui.QWidget):
    def __init__(self):
        QtGui.QWidget.__init__(self)
        self.setupUi(self)


class Ui_Form(QtGui.QWidget):
    global c2
    global c3
    global c4
    global c5
    global HighestValue
    # HighestValue = ()
    global _GlobalCableTuple
    global _GlobalCableCoreDict
    _GlobalCableCoreDict = {}
    global _CurCableID
    global CableDict
    CableDict = {}

    def __init__(self):
        QtGui.QWidget.__init__(self)
        self.setupUi(self)

    def CursorX():
        global TermX
        global DTConDes
        global DTConSrc
        global TstripX
        global PanelList
        global PanelListTagNum
        global PanelListPanelID
        global DetailList

        PanelList = []
        DetailList = []

        cursor.execute('SELECT TermID, DTermID from DetailCon WHERE DetailCon.DTermID is not Null')
        DTConSrc = cursor.fetchall()
        cursor.execute('SELECT TermID, DTermID from DetailCon WHERE DetailCon.DTermID is not Null')
        DTConDes = cursor.fetchall()
        cursor.execute('SELECT TermID, I_O from Term')
        TermX = cursor.fetchall()
        cursor.execute(
            'SELECT Panel.PanelID, Panel.TagNum, TS.TstripID, TS.TagNum, TS.AssocGID, TS.TsCon from ( Panel left join Tstrip as TS on Panel.PanelID = TS.AssocGID) ORDER BY PanelID, TS.TagNum')
        TstripX = cursor.fetchall()
        cursor.execute('SELECT Panel.PanelID, Panel.TagNum from Panel ORDER BY PanelID')
        PanelList = cursor.fetchall()

    def PopulateTermDetail(self):
        print "working..."
        cursor = cnxn.cursor()
        cursor.execute(
            "UPDATE Term SET Term.TermTag = ((cast(Tstrip.AssocGID AS VARCHAR(50)) + substring(Tstrip.TagNum,1,3)+ Term.TNum)) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID) WHERE (Tstrip.TsCon = 'F' OR Tstrip.TsCon = 'T')")
        cnxn.commit()
        cursor.execute(
            "UPDATE Term SET Term.Template = (Tstrip.TsCon) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID)")
        cnxn.commit()
        cursor.close()
        print "done"

    def OLDAssociateDuplicateTerm(self):
        # Associate Patch tray duplicate terminals
        print "working..."
        global tXSelect1
        cursor = cnxn.cursor()
        query1 = "SELECT TermID, TRow, TNum, I_O, TermTag, TNum2, Template from Term ORDER BY TermTag"
        # print query1
        cursor.execute(query1)
        tXSelect1 = cursor.fetchall()
        sum = len(tXSelect1)
        print str(sum) + " Records to be updated...."
        print "Please be patient.... Very fking patient..."
        for tXS1 in tXSelect1:
            TsCon = tXS1.Template
            series = (tXSelect1.index(tXS1))
            # print str(tXSelect1.index(tXS1)) +" / " + str(sum)


            # print tXS1.TermID, tXS1.TermTag, tXS1.TNum2, TsCon
            if TsCon == "F":
                query2a = "UPDATE Term SET Term.TNum2 = " + str(tXS1.TermID) + " WHERE (TermTag = " + "'" + str(
                    tXS1.TermTag) + "'" + " AND Template = " + "'" + "T" + "')"
                # print query2a
                cursor.execute(query2a)
                cnxn.commit()
            elif TsCon == "T":
                query2b = "UPDATE Term SET Term.TNum2 = " + str(tXS1.TermID) + " WHERE (TermTag = " + "'" + str(
                    tXS1.TermTag) + "'" + " AND Template = " + "'" + "F" + "')"
                # print query2b
                cursor.execute(query2b)
                cnxn.commit()
            else:
                pass
                # print tXS1.TermID, tXS1.TermTag, tXS1.TNum2, TsCon
        cursor.close()
        print "Congratulations! We are done!!!"

    def AssociateDuplicateTstrips(self):
        # Associate duplicate Patch trays
        print "working... start AssociateDuplicateTstrips"
        cursor = cnxn.cursor()
        queryAssociateDuplicateTstrips1a = "SELECT Tstrip.TstripID, Tstrip.TagNum, Tstrip.AssocTstripID, Tstrip.TsCon, Tstrip.AssocGID FROM Tstrip WHERE Tstrip.TsCon = 'T'"  # AND Tstrip.AssocTstripID is Null"
        cursor.execute(queryAssociateDuplicateTstrips1a)
        TstripsSelect1 = cursor.fetchall()
        sum = (len(TstripsSelect1)) * 2
        print str(sum) + " Records to be updated...."
        # print queryAssociateDuplicateTstrips1a
        time.sleep(1)
        for TSSelect1 in TstripsSelect1:
            Cur_TS_Abr1 = str(TSSelect1.TagNum[:3])
            Cur_TS_AssocGID1 = str(TSSelect1.AssocGID)
            Cur_TS_ID1 = str(TSSelect1.TstripID)
            queryAssociateDuplicateTstrips1b = "UPDATE Tstrip SET Tstrip.AssocTstripID = " + Cur_TS_ID1 + " WHERE Tstrip.TsCon = 'F' AND Tstrip.AssocGID = " + Cur_TS_AssocGID1 + " AND Tstrip.TagNum like " + "'%" + Cur_TS_Abr1 + "%'"
            # print queryAssociateDuplicateTstrips1b
            cursor.execute(queryAssociateDuplicateTstrips1b)
            cnxn.commit()
        # print "Updating TstripID: " + Cur_TS_ID1

        queryAssociateDuplicateTstrips2a = "SELECT Tstrip.TstripID, Tstrip.TagNum, Tstrip.AssocTstripID, Tstrip.TsCon, Tstrip.AssocGID FROM Tstrip WHERE Tstrip.TsCon = 'F'"  # AND Tstrip.AssocTstripID is Null"
        # print queryAssociateDuplicateTstrips2a
        time.sleep(1)
        cursor.execute(queryAssociateDuplicateTstrips2a)
        TstripsSelect2 = cursor.fetchall()
        for TSSelect2 in TstripsSelect2:
            Cur_TS_Abr2 = str(TSSelect2.TagNum[:3])
            Cur_TS_AssocGID2 = str(TSSelect2.AssocGID)
            Cur_TS_ID2 = str(TSSelect2.TstripID)
            queryAssociateDuplicateTstrips2b = "UPDATE Tstrip SET Tstrip.AssocTstripID = " + Cur_TS_ID2 + " WHERE Tstrip.TsCon = 'T' AND Tstrip.AssocGID = " + Cur_TS_AssocGID2 + " AND Tstrip.TagNum like " + "'%" + Cur_TS_Abr2 + "%'"
            # print queryAssociateDuplicateTstrips2b
            cursor.execute(queryAssociateDuplicateTstrips2b)
            cnxn.commit()
        # print "Updating TstripID:xxxxxxxxxxxxxx " + Cur_TS_ID2
        cursor.close()
        print "\nend AssociateDuplicateTstrips"

    def AssociateDuplicateTerm(self):
        # Associate duplicate Patch trays
        print "working... start AssociateDuplicateTerm"
        cursor = cnxn.cursor()
        queryAssociateDuplicateTerm1 = "SELECT Term.TNum, Term.TNum2, Term.TRow, Term.TermID, Term.TermTag, Term.TstripID, Tstrip.TagNum, Tstrip.AssocTstripID, Tstrip.TsCon, Tstrip.TstripID FROM (Tstrip left join Term on Term.TstripID = Tstrip.TstripID) WHERE Term.TNum2 IS NULL AND (Tstrip.TsCon = 'F' OR Tstrip.TsCon = 'T') AND Tstrip.AssocTstripID IS NOT Null"
        # print queryAssociateDuplicateTerm1
        time.sleep(2)
        cursor.execute(queryAssociateDuplicateTerm1)
        TermSelect1 = cursor.fetchall()
        sum = len(TermSelect1)
        print str(sum) + " Records to be updated...."
        time.sleep(2)
        for TermSel1 in TermSelect1:
            Cur_TermID = TermSel1.TermID
            Cur_Term_TNum = TermSel1.TNum
            Cur_Term_TNum2 = TermSel1.TNum2
            Cur_Term_TRow = TermSel1.TRow
            Cur_Term_TstripID = TermSel1.TstripID
            Cur_Term_TSAssocTSID = TermSel1.AssocTstripID
            queryAssociateDuplicateTerm2 = "UPDATE Term SET Term.TNum2 = " + str(
                Cur_TermID) + " FROM (Tstrip left join Term on Term.TstripID = Tstrip.TstripID) WHERE Term.TRow = " + str(
                Cur_Term_TRow) + " AND Tstrip.TstripID = " + str(Cur_Term_TSAssocTSID)
            # print queryAssociateDuplicateTerm2
            # print "Updating TermID: " + str(Cur_TermID)
            cursor.execute(queryAssociateDuplicateTerm2)
            cnxn.commit()
        cursor.close()
        print "\nend AssociateDuplicateTerm"

    def AssociateDuplicateTermSwitch(self):
        # Associate Switch duplicate terminals
        print "working..."
        cursor = cnxn.cursor()
        querySwitchUpdate1 = "SELECT Term.TermID, Term.TRow, Term.TNum, Term.I_O, Term.TermTag, Term.TNum2, Term.Template, Tstrip.TstripID, Tstrip.TsCon from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID) WHERE Tstrip.TsCon = 'N' AND Term.TNum2 is Null ORDER BY TermID"
        cursor.execute(querySwitchUpdate1)
        SwitchUpdate1 = cursor.fetchall()
        sum = len(SwitchUpdate1)
        print str(sum) + " Records to be updated...."
        time.sleep(1)
        for SwitchTerm in SwitchUpdate1:
            Cur_Switch_Term_ID = SwitchTerm.TermID
            Cur_Switch_Term_IO = SwitchTerm.I_O
            Cur_Switch_Term_TRow = SwitchTerm.TRow
            Cur_Switch_Term_Tstrip = SwitchTerm.TstripID
            print "Updating ID: " + str(Cur_Switch_Term_ID)
            if Cur_Switch_Term_TRow == 1:
                querySwitchUpdate2a = "UPDATE Term SET Term.TNum2 = " + str(
                    Cur_Switch_Term_ID) + " WHERE Term.TRow = 3 AND Term.TstripID = " + str(Cur_Switch_Term_Tstrip)
                cursor.execute(querySwitchUpdate2a)
                cnxn.commit()
            else:
                pass
            if Cur_Switch_Term_TRow == 2:
                querySwitchUpdate2b = "UPDATE Term SET Term.TNum2 = " + str(
                    Cur_Switch_Term_ID) + " WHERE Term.TRow = 4 AND Term.TstripID = " + str(Cur_Switch_Term_Tstrip)
                cursor.execute(querySwitchUpdate2b)
                cnxn.commit()
            else:
                pass
            if Cur_Switch_Term_TRow == 3:
                querySwitchUpdate2c = "UPDATE Term SET Term.TNum2 = " + str(
                    Cur_Switch_Term_ID) + " WHERE Term.TRow = 1 AND Term.TstripID = " + str(Cur_Switch_Term_Tstrip)
                cursor.execute(querySwitchUpdate2c)
                cnxn.commit()
            else:
                pass
            if Cur_Switch_Term_TRow == 4:
                querySwitchUpdate2d = "UPDATE Term SET Term.TNum2 = " + str(
                    Cur_Switch_Term_ID) + " WHERE Term.TRow = 2 AND Term.TstripID = " + str(Cur_Switch_Term_Tstrip)
                cursor.execute(querySwitchUpdate2d)
                cnxn.commit()
            else:
                pass
        print "\ndone"
        cursor.close()

    def ConnectODBC():
        global cnxn
        global cursor
        global ODBC_driver
        global ODBC_server
        global ODBC_database
        global ODBC_userID
        global userID
        global ODBC_userPassword
        global ODBC_connectionstring
        '''
        get values from ini file for the dbconnect
        '''
        ODBC_driver = 'SQL Server'
        ODBC_server = 'ZAMLRWPSQL01v\Dessoft'
        ODBC_database = '120108_Kibali_Fibre'
        # userID = getpass.getuser()
        ODBC_userID = 'dessoft_admin'
        ODBC_userPassword = 'D3ss0ft777!'

        userName = getpass.getuser()
        print "Driver: " + ODBC_driver
        print "Server Instance: " + ODBC_server
        print "Database: " + ODBC_database
        print "UserName: " + userName
        ODBC_connectionstring = 'DRIVER={' + ODBC_driver + '};SERVER=' + ODBC_server + ';DATABASE=' + ODBC_database + ';UID=' + ODBC_userID + ';PWD=' + ODBC_userPassword
        print ODBC_connectionstring
        cnxn = pyodbc.connect(ODBC_connectionstring)
        cursor = cnxn.cursor()

    def updateDestinationTermWithIO(self):
        print "start.updateDestinationTermWithIO"
        global queryUSTWI
        cursor = cnxn.cursor()
        cursor.execute('SELECT TermID, I_O from Term')
        TermXA = cursor.fetchall()
        global CurrentTermIO
        global CurTermID
        for Term in TermXA:
            if Term.I_O == '*':
                pass
            else:
                CurrentTermIO = Term.I_O
                CurTermID = Term.TermID
                for DTCS in DTConSrc:
                    if DTCS.TermID == CurTermID:
                        DesTermID = DTCS.DTermID
                        queryUDTWI = "UPDATE Term SET I_O=" + "'" + str(CurrentTermIO) + "'" + " WHERE TermID=" + str(
                            DesTermID)
                        cursor.execute(queryUDTWI)
                        cnxn.commit()
                        print queryUDTWI

                    else:
                        pass
        # cursor.close()
        print "\nend updateDestinationTermWithIO"

    def TransferDuplicateTermIO(self):
        print "start.TransferDuplicateTermIO"
        global tXSelect2
        global query2
        cursor = cnxn.cursor()
        query2 = "SELECT TermID, TRow, TNum, I_O, TermTag, TNum2, Template from Term ORDER BY TermTag"
        cursor.execute(query2)
        tXSelect2 = cursor.fetchall()
        for tXS2 in tXSelect2:
            if tXS2.I_O == "*":
                pass
            else:
                query2a = "UPDATE Term SET Term.I_O = " + "'" + str(tXS2.I_O) + "'" + " WHERE TNum2 = " + "'" + str(
                    tXS2.TermID) + "'"
                # print query2a
                cnxn.commit()
                cursor.execute(query2a)

        # cursor.close()
        print "\nend TransferDuplicateTermIO"

    def updateSourceTermWithIO(self):
        print "start.updateSourceTermWithIO"
        global queryUSTWI
        # cursor = cnxn.cursor()
        cursor.execute('SELECT TermID, I_O from Term')
        TermXB = cursor.fetchall()
        global CurrentTermIO
        global CurTermID
        for Term in TermXB:
            if Term.I_O == '*':
                pass
            else:
                CurrentTermIO = Term.I_O
                CurTermID = Term.TermID
                for DTCD in DTConDes:
                    if DTCD.DTermID == CurTermID:
                        SrcTermID = DTCD.TermID
                        queryUSTWI = "UPDATE Term SET I_O=" + "'" + str(CurrentTermIO) + "'" + " WHERE TermID=" + str(
                            SrcTermID)
                        cursor.execute(queryUSTWI)
                        # print queryUSTWI
                        cnxn.commit()
                    else:
                        pass
        # cursor.close()
        print "\nend updateSourceTermWithIO"

    def AssociateNetworks(self):
        print ".....AssociateNetworks"
        global queryAssociate1
        queryAssociate1 = "Update IIndex Set IIndex.Area = substring(II.TagNum,20,3) from ((IIndex as II left join Tstrip as TS on TS.AssocID = II.IIndexID) left join Term as T on TS.TstripID = T.TstripID) where TS.TsCon = " + "'" + "N" + "'"
        # print queryAssociate1
        cursor = cnxn.cursor()
        cursor.execute(queryAssociate1)
        cnxn.commit()
        global queryAssociate2
        queryAssociate2 = "Update Term Set Term.I_O = substring(II.Area,1,3)+ " + "'" + "a" + "'" + " from ((IIndex as II left join Tstrip as TS on TS.AssocID = II.IIndexID) left join Term as T on TS.TstripID = T.TstripID) where TS.TsCon = " + "'" + "N" + "'" + " and T.I_O = " + "'" + "*" + "'" + " and T.TRow = 1"
        # print queryAssociate2
        cursor = cnxn.cursor()
        cursor.execute(queryAssociate2)
        cnxn.commit()
        global queryAssociate3
        queryAssociate3 = "Update Term Set Term.I_O = substring(II.Area,1,3)+ " + "'" + "b" + "'" + " from ((IIndex as II left join Tstrip as TS on TS.AssocID = II.IIndexID) left join Term as T on TS.TstripID = T.TstripID) where TS.TsCon = " + "'" + "N" + "'" + " and T.I_O = " + "'" + "*" + "'" + " and T.TRow = 2"
        # print queryAssociate3
        cursor = cnxn.cursor()
        cursor.execute(queryAssociate3)
        cnxn.commit()
        cursor.close()

    def moveSwitchtoPanels(self):
        print "start moveSwitchtoPanels"
        cursor = cnxn.cursor()
        querymoveSwitchtoPanels1 = "Select CableCon.SrcCon, CableCon.SrcID, CableCon.CableID, CableCon.DesID From (CableCon inner join IIndex on CableCon.SrcID = IIndex.IIndexID) Where (CableCon.SrcCon like " + "'" + "%IIndex%" + "'" + "AND CableCon.CableID = " + "'" + "0" + "'" + " and IIndex.PanelID is NULL)"
        print querymoveSwitchtoPanels1
        cursor.execute(querymoveSwitchtoPanels1)
        CableCon = cursor.fetchall()
        # print CableCon
        for CC in CableCon:
            Cur_DesPanelID = str(CC.DesID)
            Cur_CableID = str(CC.CableID)
            Cur_Switch = str(CC.SrcID)

            if Cur_DesPanelID == 0:
                pass
            else:
                querymoveSwitchtoPanels2 = "Update IIndex Set IIndex.PanelID = " + Cur_DesPanelID + " from (CableCon left join IIndex on IIndex.IIndexID = CableCon.SrcID) Where IIndex.IIndexID = " + Cur_Switch
                # print querymoveSwitchtoPanels2
                cursor.execute(querymoveSwitchtoPanels2)
                cnxn.commit()
                querymoveSwitchtoPanels3 = "Update Tstrip Set Tstrip.AssocGID = " + Cur_DesPanelID + " from (IIndex left join Tstrip on IIndex.IIndexID = Tstrip.AssocID) Where IIndex.IIndexID = " + Cur_Switch
                # print querymoveSwitchtoPanels3
                cursor.execute(querymoveSwitchtoPanels3)
                cnxn.commit()
                querymoveSwitchtoPanels4 = "Update CableCon Set CableCon.DesID = 0 From CableCon Where SrcID = " + Cur_Switch
                # print querymoveSwitchtoPanels4
                cursor.execute(querymoveSwitchtoPanels4)
                cnxn.commit()
        cursor.close()
        print "\nend moveSwitchtoPanels"

    def updateSourceDestLoop(self):
        print "start updateSourceDestLoop"
        global CycleCount
        # global cursor
        CycleCount = 1
        # cursor = cnxn.cursor()
        while CycleCount < 10:
            self.updateDestinationTermWithIO()
            self.updateSourceTermWithIO()
            self.TransferDuplicateTermIO()
            # print CycleCount
            CycleCount += 1
        cursor.close()
        print "\nend updateSourceDestLoop"

    def ClearAllTermWithIO(self):
        print "start ClearAllTermWithIO"
        cursor = cnxn.cursor()
        query = "UPDATE Term Set Term.I_O = " + "'" + "*" + "'" + " from " + "(" + "Term left join Tstrip as TS on Term.TstripID = TS.TstripID" + ")" + " WHERE Term.I_O <> " + "'" + "*" + "'" + " AND TS.TsCon <> " + "'" + "N" + "'"
        print query
        cursor.execute(query)
        cnxn.commit()
        cursor.execute(
            "UPDATE Term SET Term.TermTag = (Null) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID)")
        print "UPDATE Term SET Term.TermTag = (Null) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID)"
        cnxn.commit()
        cursor.execute(
            "UPDATE Term SET Term.Template = (Null) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID)")
        print "UPDATE Term SET Term.Template = (Null) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID)"
        cnxn.commit()
        cursor.execute(
            "UPDATE Term SET Term.TNum2 = (Null) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID)")
        print "UPDATE Term SET Term.TNum2 = (Null) from (Tstrip left join Term on Term.TstripID = Tstrip.TstripID)"
        cnxn.commit()
        cursor.execute("UPDATE Tstrip SET Tstrip.AssocTstripID = (Null) from Tstrip")
        print "UPDATE Tstrip SET Tstrip.AssocTstripID = (Null) from Tstrip"
        cnxn.commit()
        print "done"
        cursor.close()
        print "\nend ClearAllTermWithIO"

    def Networks(self):
        nwIP4a = []
        nwIP4b = []
        nwIP4c = []
        nwIP4d = []
        nwIP4netmask = []
        nwIPnetwork = []
        nwIPnetworkDesc = []
        nwPanel = []

        NetworkFileName = 'network.csv'
        currentDirectory = os.getcwd()
        fullFilePathName = (currentDirectory + '/' + NetworkFileName)
        if not (path.isfile(fullFilePathName)):
            fileobj = open(projiniFileName, 'w')
            fileobj.close()
        else:
            with open(NetworkFileName) as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')

                for row in readCSV:
                    nwIP4a = row[0]
                    nwIP4b = row[1]
                    nwIP4c = row[2]
                    nwIP4d = row[3]
                    nwIP4netmask = row[4]
                    nwIPnetwork = row[5]
                    nwIPnetworkDesc = row[6]

            fileobj.close()

    def clearquit(self):
        print "Closing Database connections - EXIT"
        os.system('clear')
        sys.exit()

    def GetMonitorsize(self):
        global window
        global screen
        global monitors
        global nmons
        window = gtk.Window()
        screen = window.get_screen()
        monitors = []
        nmons = screen.get_n_monitors()
        print "there are %d monitors" % nmons
        for m in range(nmons):
            mg = screen.get_monitor_geometry(m)
            print "monitor %d: %d x %d" % (m, mg.width, mg.height)
            monitors.append(mg)
            print "screen size: %d x %d" % (gtk.gdk.screen_width(), gtk.gdk.screen_height())

    def paintEvent(self, e):
        qp = QtGui.QPainter()
        qp.begin(self)
        self.drawLines_in(qp)
        self.drawLines_out(qp)
        qp.end()

    def drawLines_in(self, qp):
        pen = QtGui.QPen(QtCore.Qt.darkRed, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(20, 40, 250, 40)
        pen = QtGui.QPen(QtCore.Qt.blue, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(20, 50, 250, 60)
        pen = QtGui.QPen(QtCore.Qt.green, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(20, 60, 250, 80)
        pen = QtGui.QPen(QtCore.Qt.darkYellow, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(20, 70, 250, 100)
        pen = QtGui.QPen(QtCore.Qt.gray, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(20, 80, 250, 120)

    def drawLines_out(self, qp):
        pen = QtGui.QPen(QtCore.Qt.darkRed, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(250, 40, 300, 40)
        pen = QtGui.QPen(QtCore.Qt.blue, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(250, 50, 300, 60)
        pen = QtGui.QPen(QtCore.Qt.green, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(250, 60, 300, 80)
        pen = QtGui.QPen(QtCore.Qt.darkYellow, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(250, 70, 300, 100)
        pen = QtGui.QPen(QtCore.Qt.gray, 2, QtCore.Qt.SolidLine)
        qp.setPen(pen)
        qp.drawLine(250, 80, 300, 120)

    def setupUi(self, Form):
        global c1

        def Group5A(c1, c2, c3, c4, c5):
            self.c1 = QtGui.QPushButton(Form)
            self.c1.setGeometry(QtCore.QRect(c2, c3, c4, c5))
            self.c1.setObjectName(_fromUtf8(c1))
            self.c1.setText(_translate("Form", c1, None))
            testtext = str(c1)

        # self.pushButton_2A.setText(_translate("Form", "2A", None))
        # self.pushButton_2A.clicked.connect(self.stringprint)
        Form.setObjectName(_fromUtf8("Form"))
        # Form.resize(945, 746)
        Form.setGeometry(10, 30, 1000, 1013)
        # self.textBrowser = QtGui.QTextBrowser(Form)
        # self.textBrowser.setGeometry(QtCore.QRect(240, 100, 311, 31))
        # self.textBrowser.setObjectName(_fromUtf8("textBrowser"))
        self.label = QtGui.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(150, 20, 300, 16))
        self.label.setObjectName(_fromUtf8("label"))

        self.label_2 = QtGui.QLabel(Form)
        self.label_2.setGeometry(QtCore.QRect(150, 40, 300, 16))
        self.label_2.setObjectName(_fromUtf8("label_2"))

        self.label_3 = QtGui.QLabel(Form)
        self.label_3.setGeometry(QtCore.QRect(50, 20, 200, 16))
        self.label_3.setObjectName(_fromUtf8("label_2"))

        self.label_4 = QtGui.QLabel(Form)
        self.label_4.setGeometry(QtCore.QRect(50, 40, 200, 16))
        self.label_4.setObjectName(_fromUtf8("label_2"))
        '''
        #self.textBrowser_2 = QtGui.QTextBrowser(Form)
        #self.textBrowser_2.setGeometry(QtCore.QRect(240, 140, 311, 31))
        #self.textBrowser_2.setObjectName(_fromUtf8("textBrowser_2"))
        self.checkBox = QtGui.QCheckBox(Form)
        self.checkBox.setGeometry(QtCore.QRect(240, 190, 70, 17))
        self.checkBox.setObjectName(_fromUtf8("checkBox"))
        self.pushButton = QtGui.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(69, 237, 175, 23))
        self.pushButton.setObjectName(_fromUtf8("pushButton"))
        '''
        self.pushButton_2 = QtGui.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(50, 440, 300, 23))
        self.pushButton_2.setObjectName(_fromUtf8("pushButton_2"))

        self.pushButton_3 = QtGui.QPushButton(Form)
        self.pushButton_3.setGeometry(QtCore.QRect(50, 380, 300, 23))
        self.pushButton_3.setObjectName(_fromUtf8("pushButton_3"))

        self.pushButton_4 = QtGui.QPushButton(Form)
        self.pushButton_4.setGeometry(QtCore.QRect(50, 410, 300, 23))
        self.pushButton_4.setObjectName(_fromUtf8("pushButton_4"))

        self.pushButton_5 = QtGui.QPushButton(Form)
        self.pushButton_5.setGeometry(QtCore.QRect(50, 350, 300, 23))
        self.pushButton_5.setObjectName(_fromUtf8("pushButton_5"))
        self.pushButton_5.setEnabled(True)
        # self.pushButton_5.hide()

        self.pushButton_6 = QtGui.QPushButton(Form)
        self.pushButton_6.setGeometry(QtCore.QRect(200, 60, 175, 23))
        self.pushButton_6.setObjectName(_fromUtf8("pushButton_6"))

        self.pushButton_7 = QtGui.QPushButton(Form)
        self.pushButton_7.setGeometry(QtCore.QRect(50, 110, 300, 23))
        self.pushButton_7.setObjectName(_fromUtf8("pushButton_7"))

        self.pushButton_8 = QtGui.QPushButton(Form)
        self.pushButton_8.setGeometry(QtCore.QRect(50, 170, 300, 23))
        self.pushButton_8.setObjectName(_fromUtf8("pushButton_8"))

        self.pushButton_9 = QtGui.QPushButton(Form)
        self.pushButton_9.setGeometry(QtCore.QRect(50, 140, 300, 23))
        self.pushButton_9.setObjectName(_fromUtf8("pushButton_9"))

        self.pushButton_10 = QtGui.QPushButton(Form)
        self.pushButton_10.setGeometry(QtCore.QRect(50, 200, 300, 23))
        self.pushButton_10.setObjectName(_fromUtf8("pushButton_10"))

        self.pushButton_11 = QtGui.QPushButton(Form)
        self.pushButton_11.setGeometry(QtCore.QRect(50, 230, 300, 23))
        self.pushButton_11.setObjectName(_fromUtf8("pushButton_11"))

        self.pushButton_12 = QtGui.QPushButton(Form)
        self.pushButton_12.setGeometry(QtCore.QRect(50, 320, 300, 23))
        self.pushButton_12.setObjectName(_fromUtf8("pushButton_12"))

        self.pushButton_13 = QtGui.QPushButton(Form)
        self.pushButton_13.setGeometry(QtCore.QRect(50, 290, 300, 23))
        self.pushButton_13.setObjectName(_fromUtf8("pushButton_13"))

        self.pushButton_14 = QtGui.QPushButton(Form)
        self.pushButton_14.setGeometry(QtCore.QRect(50, 80, 300, 23))
        self.pushButton_14.setObjectName(_fromUtf8("pushButton_14"))

        self.pushButton_15 = QtGui.QPushButton(Form)
        self.pushButton_15.setGeometry(QtCore.QRect(50, 260, 300, 23))
        self.pushButton_15.setObjectName(_fromUtf8("pushButton_15"))

        self.pushButton_16 = QtGui.QPushButton(Form)
        self.pushButton_16.setGeometry(QtCore.QRect(800, 690, 75, 23))
        self.pushButton_16.setObjectName(_fromUtf8("pushButton_16"))

        self.pushButton_17 = QtGui.QPushButton(Form)
        self.pushButton_17.setGeometry(QtCore.QRect(700, 10, 200, 23))
        self.pushButton_17.setObjectName(_fromUtf8("pushButton_17"))

        self.pushButton_18 = QtGui.QPushButton(Form)
        self.pushButton_18.setGeometry(QtCore.QRect(700, 40, 200, 23))
        self.pushButton_18.setObjectName(_fromUtf8("pushButton_18"))

        self.pushButton_19 = QtGui.QPushButton(Form)
        self.pushButton_19.setGeometry(QtCore.QRect(700, 70, 200, 23))
        self.pushButton_19.setObjectName(_fromUtf8("pushButton_19"))

        self.pushButton_20 = QtGui.QPushButton(Form)
        self.pushButton_20.setGeometry(QtCore.QRect(700, 100, 200, 23))
        self.pushButton_20.setObjectName(_fromUtf8("pushButton_20"))

        self.pushButton_21 = QtGui.QPushButton(Form)
        self.pushButton_21.setGeometry(QtCore.QRect(700, 130, 200, 23))
        self.pushButton_21.setObjectName(_fromUtf8("pushButton_21"))

        self.pushButton_22 = QtGui.QPushButton(Form)
        self.pushButton_22.setGeometry(QtCore.QRect(700, 160, 200, 23))
        self.pushButton_22.setObjectName(_fromUtf8("pushButton_22"))

        self.pushButton_23 = QtGui.QPushButton(Form)
        self.pushButton_23.setGeometry(QtCore.QRect(700, 190, 200, 23))
        self.pushButton_23.setObjectName(_fromUtf8("pushButton_23"))

        self.pushButton_24 = QtGui.QPushButton(Form)
        self.pushButton_24.setGeometry(QtCore.QRect(700, 220, 200, 23))
        self.pushButton_24.setObjectName(_fromUtf8("pushButton_24"))

        self.pushButton_25 = QtGui.QPushButton(Form)
        self.pushButton_25.setGeometry(QtCore.QRect(700, 280, 200, 23))
        self.pushButton_25.setObjectName(_fromUtf8("pushButton_25"))

        Group5A('pushButton_2A', 500, 10, 100, 23)
        Group5A('pushButton_2X', 500, 40, 100, 23)
        Group5A('pushButton_2Z', 500, 70, 100, 23)
        # self.c1.clicked.connect(self.stringprint)

        # self.pushButton_2A.setEnabled(True)
        # self.pushButton_2A.hide()

        self.pushButton_2B = QtGui.QPushButton(Form)
        self.pushButton_2B.setGeometry(QtCore.QRect(400, 40, 50, 23))
        self.pushButton_2B.setObjectName(_fromUtf8("pushButton_2B"))

        self.pushButton_2C = QtGui.QPushButton(Form)
        self.pushButton_2C.setGeometry(QtCore.QRect(400, 70, 50, 23))
        self.pushButton_2C.setObjectName(_fromUtf8("pushButton_2C"))

        self.pushButton_2D = QtGui.QPushButton(Form)
        self.pushButton_2D.setGeometry(QtCore.QRect(400, 100, 50, 23))
        self.pushButton_2D.setObjectName(_fromUtf8("pushButton_2D"))

        self.pushButton_2E = QtGui.QPushButton(Form)
        self.pushButton_2E.setGeometry(QtCore.QRect(400, 130, 50, 23))
        self.pushButton_2E.setObjectName(_fromUtf8("pushButton_2E"))

        self.pushButton_2F = QtGui.QPushButton(Form)
        self.pushButton_2F.setGeometry(QtCore.QRect(400, 160, 50, 23))
        self.pushButton_2F.setObjectName(_fromUtf8("pushButton_2F"))

        self.pushButton_2G = QtGui.QPushButton(Form)
        self.pushButton_2G.setGeometry(QtCore.QRect(400, 190, 50, 23))
        self.pushButton_2G.setObjectName(_fromUtf8("pushButton_2G"))

        self.pushButton_2H = QtGui.QPushButton(Form)
        self.pushButton_2H.setGeometry(QtCore.QRect(400, 220, 50, 23))
        self.pushButton_2H.setObjectName(_fromUtf8("pushButton_2H"))

        self.pushButton_2I = QtGui.QPushButton(Form)
        self.pushButton_2I.setGeometry(QtCore.QRect(400, 250, 50, 23))
        self.pushButton_2I.setObjectName(_fromUtf8("pushButton_2I"))

        self.pushButton_2J = QtGui.QPushButton(Form)
        self.pushButton_2J.setGeometry(QtCore.QRect(400, 280, 50, 23))
        self.pushButton_2J.setObjectName(_fromUtf8("pushButton_2J"))

        self.pushButton_2K = QtGui.QPushButton(Form)
        self.pushButton_2K.setGeometry(QtCore.QRect(400, 310, 50, 23))
        self.pushButton_2K.setObjectName(_fromUtf8("pushButton_2K"))

        self.pushButton_2L = QtGui.QPushButton(Form)
        self.pushButton_2L.setGeometry(QtCore.QRect(400, 340, 50, 23))
        self.pushButton_2L.setObjectName(_fromUtf8("pushButton_2L"))

        self.pushButton_2M = QtGui.QPushButton(Form)
        self.pushButton_2M.setGeometry(QtCore.QRect(400, 370, 50, 23))
        self.pushButton_2M.setObjectName(_fromUtf8("pushButton_2M"))

        self.pushButton_2N = QtGui.QPushButton(Form)
        self.pushButton_2N.setGeometry(QtCore.QRect(400, 400, 50, 23))
        self.pushButton_2N.setObjectName(_fromUtf8("pushButton_2N"))

        self.pushButton_2O = QtGui.QPushButton(Form)
        self.pushButton_2O.setGeometry(QtCore.QRect(400, 430, 50, 23))
        self.pushButton_2O.setObjectName(_fromUtf8("pushButton_2O"))

        self.pushButton_2P = QtGui.QPushButton(Form)
        self.pushButton_2P.setGeometry(QtCore.QRect(400, 460, 50, 23))
        self.pushButton_2P.setObjectName(_fromUtf8("pushButton_2P"))

        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)

    def retranslateUi(self, Form):
        Form.setWindowTitle(_translate("Form", ODBC_database, None))
        self.label.setText(_translate("Form", ODBC_server, None))
        self.label_2.setText(_translate("Form", ODBC_database, None))
        self.label_3.setText(_translate("Form", "Server Instance:", None))
        self.label_4.setText(_translate("Form", "Database:", None))
        '''
        self.checkBox.setText(_translate("Form", "CheckBox", None))
        self.pushButton.setText(_translate("Form", "PushButton1", None))
        '''

        self.pushButton_2.setText(_translate("Form", "PopulateNetworkTable", None))
        self.pushButton_2.clicked.connect(self.PopulateNetworkTable)

        self.pushButton_3.setText(_translate("Form", "Update Dwg Descriptions", None))
        self.pushButton_3.clicked.connect(self.UpdateDwgDescriptions)

        self.pushButton_4.setText(_translate("Form", "Update Drawing Template per Revision", None))
        self.pushButton_4.clicked.connect(self.UpdateDrawingTemplateRevision)

        self.pushButton_5.setText(_translate("Form", "Associate Cables with Drawings && Templates", None))
        self.pushButton_5.clicked.connect(self.AssociateCableswithDrawingsTemplates)

        self.pushButton_6.setText(_translate("Form", "UpdateTableStructure", None))
        self.pushButton_6.clicked.connect(self.UpdateTableStructure)

        self.pushButton_7.setText(_translate("Form", "Update CableTags", None))
        self.pushButton_7.clicked.connect(self.UpdateCableTags)

        self.pushButton_8.setText(_translate("Form", "Associate Duplicate TStrips", None))
        self.pushButton_8.clicked.connect(self.AssociateDuplicateTstrips)

        self.pushButton_9.setText(_translate("Form", "Populate Term Detail", None))
        self.pushButton_9.clicked.connect(self.PopulateTermDetail)

        self.pushButton_10.setText(_translate("Form", "Associate Duplicate Terminals", None))
        self.pushButton_10.clicked.connect(self.AssociateDuplicateTerm)

        self.pushButton_11.setText(_translate("Form", "Move Switch to Panels", None))
        self.pushButton_11.clicked.connect(self.moveSwitchtoPanels)

        self.pushButton_12.setText(_translate("Form", "Update Terminals With Network", None))
        self.pushButton_12.clicked.connect(self.updateSourceDestLoop)

        self.pushButton_13.setText(_translate("Form", "Associate Duplicate Term Switch", None))
        self.pushButton_13.clicked.connect(self.AssociateDuplicateTermSwitch)

        self.pushButton_14.setText(_translate("Form", "Clear Network Allocations", None))
        self.pushButton_14.clicked.connect(self.ClearAllTermWithIO)

        self.pushButton_15.setText(_translate("Form", "Associate Switches with Network", None))
        self.pushButton_15.clicked.connect(self.AssociateNetworks)

        self.pushButton_16.setText(_translate("Form", "Quit", None))
        self.pushButton_16.clicked.connect(self.clearquit)

        self.pushButton_17.setText(_translate("Form", "Populate Midcoupler Table", None))
        self.pushButton_17.clicked.connect(self.cPopulateMidcouplerTable)

        self.pushButton_18.setText(_translate("Form", "Populate Tray Table", None))
        self.pushButton_18.clicked.connect(self.cPopulateTrayTable)

        self.pushButton_19.setText(_translate("Form", "Populate Cable Table", None))
        self.pushButton_19.clicked.connect(self.cPopulateCableTable)

        self.pushButton_20.setText(_translate("Form", "Populate Panel Table", None))
        self.pushButton_20.clicked.connect(self.cPopulatePanelTable)

        self.pushButton_21.setText(_translate("Form", "Populate Core Table", None))
        self.pushButton_21.clicked.connect(self.cPopulateCoreTable)

        self.pushButton_22.setText(_translate("Form", "Populate DetailCon Table per Cable", None))
        self.pushButton_22.clicked.connect(self.cPopulateDetailCon1)

        # self.pushButton_23.setText(_translate("Form", "Create Patch Termination drawing Numbers", None))
        # self.pushButton_23.clicked.connect(self.cCreatePatchDocuments)

        self.pushButton_24.setText(_translate("Form", "cPopulateDrawingNumbers_Patch", None))
        self.pushButton_24.clicked.connect(self.cPopulateDrawingNumbers_Patch)

        self.pushButton_25.setText(_translate("Form", "ReadTableStructure", None))
        self.pushButton_25.clicked.connect(self.ReadTableStructure)

        # self.c1.setText(_translate("Form", c1, None))
        # self.c1.clicked.connect(self.stringprint)

        # self.(c1).setText(_translate("Form", "pushButton_2X", None))
        # self.pushButton_2A.clicked.connect(self.stringprint)
        # self.c1.setText(_translate("Form", c1, None))
        ###self.c1.clicked.connect(self.stringprint)

        # self.pushButton_2A.setText(_translate("Form", "2A", None))
        # self.pushButton_2A.clicked.connect(self.stringprint)

        # self.pushButton_2A.setText(_translate("Form", "2A", None))
        # self.pushButton_2A.clicked.connect(self.clearquit)

        self.pushButton_2B.setText(_translate("Form", "2B", None))
        # self.pushButton_2B.clicked.connect(self.clearquit)

        self.pushButton_2C.setText(_translate("Form", "2C", None))
        # self.pushButton_2C.clicked.connect(self.clearquit)

        self.pushButton_2D.setText(_translate("Form", "2D", None))
        # self.pushButton_2D.clicked.connect(self.clearquit)

        self.pushButton_2E.setText(_translate("Form", "2E", None))
        # self.pushButton_2E.clicked.connect(self.clearquit)

        self.pushButton_2F.setText(_translate("Form", "2F", None))
        # self.pushButton_2F.clicked.connect(self.clearquit)

        self.pushButton_2G.setText(_translate("Form", "2G", None))
        # self.pushButton_2G.clicked.connect(self.clearquit)

        self.pushButton_2H.setText(_translate("Form", "2H", None))
        # self.pushButton_2H.clicked.connect(self.clearquit)

        self.pushButton_2I.setText(_translate("Form", "2I", None))
        # self.pushButton_2I.clicked.connect(self.clearquit)

        self.pushButton_2J.setText(_translate("Form", "2J", None))
        # self.pushButton_2J.clicked.connect(self.clearquit)

        self.pushButton_2K.setText(_translate("Form", "2K", None))
        # self.pushButton_2K.clicked.connect(self.clearquit)

        self.pushButton_2L.setText(_translate("Form", "2L", None))
        # self.pushButton_2L.clicked.connect(self.clearquit)

        self.pushButton_2M.setText(_translate("Form", "2M", None))
        # self.pushButton_2M.clicked.connect(self.clearquit)

        self.pushButton_2N.setText(_translate("Form", "2N", None))
        # self.pushButton_2N.clicked.connect(self.clearquit)

        self.pushButton_2O.setText(_translate("Form", "2O", None))
        # self.pushButton_2O.clicked.connect(self.clearquit)

        self.pushButton_2P.setText(_translate("Form", "2P", None))

    def stringprint(self, testtext):
        print testtext

    def CreateTSbyPanelLists(self):
        cursor = cnxn.cursor()
        global CurPanelTagNum
        global CurPanelID
        for PanX in PanelList:
            if Term.I_O == '*':
                pass
            else:
                CurrentTermIO = Term.I_O
                CurTermID = Term.TermID
                for DTCS in DTConSrc:
                    if DTCS.TermID == CurTermID:
                        DesTermID = DTCS.DTermID
                        query = "UPDATE Term SET I_O=" + "'" + str(CurrentTermIO) + "'" + " WHERE TermID=" + str(
                            DesTermID)
                        print query
                        cursor.execute(query)
                        cnxn.commit()
                    else:
                        pass
        cursor.close()

    def UpdateCableTags(self):
        print "start UpdateCableTags"
        cursor = cnxn.cursor()
        Q1_UpdateCableTags = "UPDATE Cable SET Cable.CableTagSource = Cable.CableTag1+'.'+STS.Display from (((((((Cable inner join CableCon ON CableCon.CableID = Cable.CableID) inner join Core ON CableCon.CableID  = Core.CableID) inner join DetailCon on DetailCon.CoreID = Core.CoreID) inner join Term AS STERM ON STERM.TermID = DetailCon.TermID) inner join Term AS DTERM ON DTERM.TermID = DetailCon.DTermID) inner join Tstrip as STS on STS.TstripID = STERM.TstripID) inner join Tstrip as DTS on DTS.TstripID = DTERM.TstripID) where Core.Ord = 1 and Cable.CType = 'BUS' and CableCon.SrcCon like '%Panel%'"
        cursor.execute(Q1_UpdateCableTags)
        cnxn.commit()
        Q2_UpdateCableTags = "UPDATE Cable SET Cable.CableTagDestination = Cable.CableTag2+'.'+DTS.Display from (((((((Cable inner join CableCon ON CableCon.CableID = Cable.CableID) inner join Core ON CableCon.CableID  = Core.CableID) inner join DetailCon on DetailCon.CoreID = Core.CoreID) inner join Term AS STERM ON STERM.TermID = DetailCon.TermID) inner join Term AS DTERM ON DTERM.TermID = DetailCon.DTermID) inner join Tstrip as STS on STS.TstripID = STERM.TstripID) inner join Tstrip as DTS on DTS.TstripID = DTERM.TstripID) where Core.Ord = 1 and Cable.CType = 'BUS' and CableCon.SrcCon like '%Panel%'"
        cursor.execute(Q2_UpdateCableTags)
        cnxn.commit()
        cursor.close()
        print "\nend UpdateCableTags"

    def SelectCores(_CurCableID):
        CurCores = []
        cursor = cnxn.cursor()
        Q1_SelectCores = "SELECT Core.CoreID FROM (Core LEFT JOIN Cable ON Core.CableID = Cable.CableID) WHERE Cable.CableID = " + "'" + str(
            _CurCableID) + "'"
        cursor.execute(Q1_SelectCores)
        Q1_SelectCores_result = cursor.fetchall()
        cursor.close()

        CurCores = Q1_SelectCores_result
        print Q1_SelectCores_result

    def UpdateDrawingTemplateRevision(self):
        print "start UpdateDrawingTemplateRevision"
        cursor = cnxn.cursor()
        Q1_DrawingTemplateRevision = "Update Document Set Document.Template = Document.MainTemplate + '_.dwg' from Document where Rev like '%B%'"
        cursor.execute(Q1_DrawingTemplateRevision)
        cnxn.commit()
        Q2_DrawingTemplateRevision = "Update Document Set Document.Template = Document.MainTemplate + '.dwg' from Document where Rev not like '%B%'"
        cursor.execute(Q2_DrawingTemplateRevision)
        cnxn.commit()
        cursor.close()
        print "\nend UpdateDrawingTemplateRevision"

    def createCableDict(self):

        # Create list with all cables, and associated panels
        q2_cables = "SELECT Cable.CableID, Cable.OtherDocs, Cable.EstLength, Cable.CableTemplate, Cable.TagNum, Cable.CableTag1, Cable.CableTag2, CableCon.CableID, CableCon.SrcID FROM Cable left join CableCon on CableCon.CableID = Cable.CableID"
        cursor.execute(q2_cables)
        q2_cables_result = cursor.fetchall()

        # Create list with all cores per cables
        q3_cores = "SELECT Core.CoreID, Core.Core, Core.Network, Core.Ord, Core.CableID FROM Core"
        cursor.execute(q3_cores)
        q3_cores_result = cursor.fetchall()

        # Create list with all Tstrips per panel
        q4_tstrip = "SELECT Tstrip.TstripID, Tstrip.TsCon, Tstrip.AssocGID, Tstrip.Display, Tstrip.AssocTstripID FROM Tstrip"
        cursor.execute(q4_tstrip)
        q4_tstrip_result = cursor.fetchall()

        # create a list with all detailconnected cores with cable, source tstrip, terminal:
        q5_detailcon = "SELECT DetailCon.CableID, DetailCon.CoreID, DetailCon.TermID, DetailCon.TstripID, DetailCon.SCon from DetailCon where DetailCon.CableID <> ''"
        cursor.execute(q5_detailcon)
        q5_detailcon_result = cursor.fetchall()

        # Create list with all Terms per Tstrip
        q6_term = "SELECT Term.I_O, Term.TNum, term.TRow, term.TstripID, term.TermID from term"
        cursor.execute(q6_term)
        q6_term_result = cursor.fetchall()

        print "------q1_panels_result-------"
        print q1_panels_result
        print "------q2_cables_result-------"
        print q2_cables_result
        print "------q3_cores_result-------"
        print q3_cores_result
        print "------q4_tstrip_result-------"
        print q4_tstrip_result
        print "------q5_detailcon_result-------"
        print q5_detailcon_result
        print "------q6_term_result-------"
        print q6_term_result

        cursor.close()

    def UpdateDwgDescriptions(self):
        print "Working..."
        print "Update Dwg Descriptions"
        cursor = cnxn.cursor()
        Q1_UpdateDwgDescriptions = "Update Cable Set Cable.OtherDocs = '120108-DWG-07GB-108-' + SUBSTRING(Cable.Notes,1,7) from Cable where cable.OtherDocs is not NULL"
        cursor.execute(Q1_UpdateDwgDescriptions)
        cnxn.commit()
        Q2_UpdateDwgDescriptions = "Update Document Set Document.Description = 'Data Communications - Source: ' + Cable.CableTagSource + ' - Destination: ' + Cable.CableTagDestination + ' - Splice Termination Diagram' from Document inner join Cable ON Cable.OtherDocs = Document.FileName"
        cursor.execute(Q2_UpdateDwgDescriptions)
        cnxn.commit()
        cursor.close()
        print "Done"

    def cPopulateMidcouplerTable(self):
        print "Working..."
        print "Populate Midcoupler Table..."
        global HighestValue
        HighestValue = ()
        cursor = cnxn.cursor()
        Q1_cPopulateMidcouplerTable = "select * from Term inner join Tstrip on Tstrip.TstripID = Term.TstripID where ((Tstrip.TsCon = " + "'" + "F" + "') AND (Term.Populated is NULL))"
        cursor.execute(Q1_cPopulateMidcouplerTable)
        Q1_result = cursor.fetchall()
        for MC in Q1_result:
            if MC.TermID == None:
                Cur_MidcouplerID = 'NULL'
            else:
                Cur_MidcouplerID = MC.TermID
            if MC.TRow == None:
                Cur_Ord = 'NULL'
            else:
                Cur_Ord = MC.TRow
            if MC.TsCon == None:
                Cur_TSCon = 'NULL'
            else:
                Cur_TSCon = "'" + str(MC.TsCon) + "'"
            if MC.TNum == None:
                Cur_TagNum = 'NULL'
            else:
                Cur_TagNum = "'" + str(MC.TNum) + "'"
            if MC.TstripID == None:
                Cur_AssocID = 'NULL'
            else:
                Cur_AssocID = MC.TstripID
            Cur_Signal = 'NULL'
            Cur_Notes = 'NULL'
            Cur_DwgSeq = 'NULL'
            Cur_Sht = 'NULL'
            Cur_X = 'NULL'
            Cur_Y = 'NULL'
            Q2_cPopulateMidcouplerTable = "INSERT INTO cMidcoupler (MidcouplerID, TagNum, AssocID, Ord, Signal, Notes, Cur_DwgSeq, Cur_Sht, Cur_X, Cur_Y) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)" % (
            Cur_MidcouplerID, Cur_TagNum, Cur_AssocID, Cur_Ord, Cur_Signal, Cur_Notes, Cur_DwgSeq, Cur_Sht, Cur_X,
            Cur_Y)
            cursor.execute(Q2_cPopulateMidcouplerTable)
            cnxn.commit()
            Q3_cPopulateMidcouplerTable = "Update Term SET Term.Populated = 1 WHERE Term.TermID = %s" % (
            Cur_MidcouplerID)
            cursor.execute(Q3_cPopulateMidcouplerTable)
            cnxn.commit()
            Q_MCHigehestValue = "select cLastUsedValues.HighValue from cLastUsedValues"
            cursor.execute(Q_MCHigehestValue)
            HighestValue = cursor.fetchall()
            if Cur_MidcouplerID > HighestValue:
                Q_Replace_MCHV = "Update cLastUsedValues SET cLastUsedValues.HighValue = %s" % (Cur_MidcouplerID)
                cursor.execute(Q_Replace_MCHV)
                cnxn.commit()
            else:
                pass
        cursor.close()
        print "Done"

    def cPopulateTrayTable(self):
        print "Working..."
        print "Populate Tray Table..."
        global HighestValue
        HighestValue = ()
        cursor = cnxn.cursor()
        Q1_cPopulateTrayTable = "select * from Tstrip where ((Tstrip.TsCon = " + "'" + "F" + "') AND (Tstrip.Populated is NULL))"
        cursor.execute(Q1_cPopulateTrayTable)
        Q1_result = cursor.fetchall()
        for PTray in Q1_result:
            if PTray.TstripID == None:
                Cur_TrayID = 'NULL'
            else:
                Cur_TrayID = PTray.TstripID
            if PTray.Ord == None:
                Cur_Ord = 'NULL'
            else:
                Cur_Ord = PTray.Ord
            if PTray.TagNum == None:
                Cur_TagNum = 'NULL'
            else:
                Cur_TagNum = "'" + str(PTray.TagNum) + "'"
            if PTray.TagRule == None:
                Cur_TagRule = 'NULL'
            else:
                Cur_TagRule = "'" + str(PTray.TagRule) + "'"
            if PTray.AssocGID == None:
                Cur_AssocID = 'NULL'
            else:
                Cur_AssocID = PTray.AssocGID
            if PTray.Display == None:
                Cur_Display = 'NULL'
            else:
                Cur_Display = "'" + str(PTray.Display) + "'"
            if PTray.Notes == None:
                Cur_Notes = 'NULL'
            else:
                Cur_Notes = "'" + str(PTray.Notes) + "'"
            if PTray.OtherDocs == None:
                Cur_OtherDocs = 'NULL'
            else:
                Cur_OtherDocs = "'" + str(PTray.OtherDocs) + "'"
            Q2_cPopulateTrayTable = "INSERT INTO cTray (TrayID, TagNum, TagRule, Ord, AssocID, Notes, Display, OtherDocs) VALUES (%s,%s,%s,%s,%s,%s,%s,%s)" % (
            Cur_TrayID, Cur_TagNum, Cur_TagRule, Cur_Ord, Cur_AssocID, Cur_Notes, Cur_Display, Cur_OtherDocs)
            # print Q2_cPopulateTrayTable
            cursor.execute(Q2_cPopulateTrayTable)
            cnxn.commit()
            Q3_cPopulateTrayTable = "Update Tstrip SET Tstrip.Populated = 1 WHERE Tstrip.TstripID = %s" % (Cur_TrayID)
            cursor.execute(Q3_cPopulateTrayTable)
            cnxn.commit()
            Q_MCHigehestValue = "select cLastUsedValues.HighValue from cLastUsedValues"
            cursor.execute(Q_MCHigehestValue)
            HighestValue = cursor.fetchall()
            if Cur_TrayID > HighestValue:
                Q_Replace_MCHV = "Update cLastUsedValues SET cLastUsedValues.HighValue = %s" % (Cur_TrayID)
                cursor.execute(Q_Replace_MCHV)
                cnxn.commit()
            else:
                pass
        cursor.close()
        print "Done"

    def cPopulateCableTable(self):
        print "Working..."
        print "Populate Cable Table..."
        global HighestValue
        HighestValue = ()
        cursor = cnxn.cursor()
        Q1_cPopulateCableTable = "select * from Cable WHERE (Cable.Populated is NULL)"
        cursor.execute(Q1_cPopulateCableTable)
        Q1_result = cursor.fetchall()
        for CAB in Q1_result:
            if CAB.CableID == None:
                Cur_CableID = 'NULL'
            else:
                Cur_CableID = CAB.CableID
            if CAB.TagNum == None:
                Cur_TagNum = 'NULL'
            else:
                Cur_TagNum = "'" + str(CAB.TagNum) + "'"
            if CAB.TagRule == None:
                Cur_TagRule = 'NULL'
            else:
                Cur_TagRule = "'" + str(CAB.TagRule) + "'"
            if CAB.CableSTD == None:
                Cur_CabSTD = 'NULL'
            else:
                Cur_CabSTD = "'" + str(CAB.CableSTD) + "'"
            if CAB.CableTagSource == None:
                Cur_CabTagSRC = 'NULL'
            else:
                Cur_CabTagSRC = "'" + str(CAB.CableTagSource) + "'"
            if CAB.CableTagDestination == None:
                Cur_CabTagDES = 'NULL'
            else:
                Cur_CabTagDES = "'" + str(CAB.CableTagDestination) + "'"
            if CAB.Area == None:
                Cur_Area = 'NULL'
            else:
                Cur_Area = "'" + str(CAB.Area) + "'"
            if CAB.Func == None:
                Cur_Func = 'NULL'
            else:
                Cur_Func = "'" + str(CAB.Func) + "'"
            if CAB.Num == None:
                Cur_Num = 'NULL'
            else:
                Cur_Num = "'" + str(CAB.Num) + "'"
            if CAB.Suffix == None:
                Cur_Suffix = 'NULL'
            else:
                Cur_Suffix = "'" + str(CAB.Suffix) + "'"
            if CAB.OtherDocs == None:
                Cur_OtherDocs = 'NULL'
            else:
                Cur_OtherDocs = "'" + str(CAB.OtherDocs) + "'"
            if CAB.OtherDocs2 == None:
                Cur_OtherDocs2 = 'NULL'
            else:
                Cur_OtherDocs2 = "'" + str(CAB.OtherDocs2) + "'"
            if CAB.Length == None:
                Cur_Length = 'NULL'
            else:
                Cur_Length = CAB.Length
            if CAB.EstLength == None:
                Cur_EstLength = 'NULL'
            else:
                Cur_EstLength = CAB.EstLength
            if CAB.Notes == None:
                Cur_Notes = 'NULL'
            else:
                Cur_Notes = "'" + str(CAB.Notes) + "'"
            if CAB.Rev == None:
                Cur_Rev = 'NULL'
            else:
                Cur_Rev = "'" + str(CAB.Rev) + "'"
            Q2_cPopulateCableTable = "INSERT INTO cCable (CableID, TagNum, TagRule, CableSTD, CableTagSource, CableTagDestination, Area, Func, Num, Suffix, OtherDocs, OtherDocs2, Length, EstLength, Notes, Rev) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)" % (
            Cur_CableID, Cur_TagNum, Cur_TagRule, Cur_CabSTD, Cur_CabTagSRC, Cur_CabTagDES, Cur_Area, Cur_Func, Cur_Num,
            Cur_Suffix, Cur_OtherDocs, Cur_OtherDocs2, Cur_Length, Cur_EstLength, Cur_Notes, Cur_Rev)
            cursor.execute(Q2_cPopulateCableTable)
            cnxn.commit()
            Q3_cPopulateCableTable = "Update Cable SET Cable.Populated = 1 WHERE Cable.CableID = %s" % (Cur_CableID)
            cursor.execute(Q3_cPopulateCableTable)
            cnxn.commit()
            Q_MCHigehestValue = "select cLastUsedValues.HighValue from cLastUsedValues"
            cursor.execute(Q_MCHigehestValue)
            HighestValue = cursor.fetchall()
            if Cur_CableID > HighestValue:
                Q_Replace_MCHV = "Update cLastUsedValues SET cLastUsedValues.HighValue = %s" % (Cur_CableID)
                cursor.execute(Q_Replace_MCHV)
                cnxn.commit()
            else:
                pass
        cursor.close()
        print "Done"

    def cPopulatePanelTable(self):
        print "Working..."
        print "Populate Panel Table..."
        global HighestValue
        HighestValue = ()
        cursor = cnxn.cursor()
        Q1_cPopulatePanelTable = "select * from Panel WHERE (Panel.Populated is NULL)"
        cursor.execute(Q1_cPopulatePanelTable)
        Q1_result = cursor.fetchall()
        for PAN in Q1_result:
            if PAN.TagNum == None:
                Cur_TagNum = 'NULL'
            else:
                Cur_TagNum = "'" + str(PAN.TagNum) + "'"
            if PAN.Description == None:
                Cur_Description = 'NULL'
            else:
                Cur_Description = "'" + str(PAN.Description) + "'"
            if PAN.Ord == None:
                Cur_Ord = 'NULL'
            else:
                Cur_Ord = PAN.Ord
            if PAN.Area == None:
                Cur_Area = 'NULL'
            else:
                Cur_Area = "'" + str(PAN.Area) + "'"
            if PAN.Func == None:
                Cur_Func = 'NULL'
            else:
                Cur_Func = "'" + str(PAN.Func) + "'"
            if PAN.Num == None:
                Cur_Num = 'NULL'
            else:
                Cur_Num = "'" + str(PAN.Num) + "'"
            if PAN.OtherDocs == None:
                Cur_OtherDocs = 'NULL'
            else:
                Cur_OtherDocs = "'" + str(PAN.OtherDocs) + "'"
            if PAN.TagRule == None:
                Cur_TagRule = 'NULL'
            else:
                Cur_TagRule = "'" + str(PAN.TagRule) + "'"
            if PAN.Notes == None:
                Cur_Notes = 'NULL'
            else:
                Cur_Notes = "'" + str(PAN.Notes) + "'"
            if PAN.Rev == None:
                Cur_Rev = 'NULL'
            else:
                Cur_Rev = "'" + str(PAN.Rev) + "'"
            if PAN.PanelID == None:
                Cur_PanelID = 'NULL'
            else:
                Cur_PanelID = PAN.PanelID
            Q2_cPopulatePanelTable = "INSERT INTO cPanel (PanelID, TagNum, Ord, Description, Area, Func, Num, OtherDocs, TagRule, Notes, Rev) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)" % (
            Cur_PanelID, Cur_TagNum, Cur_Ord, Cur_Description, Cur_Area, Cur_Func, Cur_Num, Cur_OtherDocs, Cur_TagRule,
            Cur_Notes, Cur_Rev)
            cursor.execute(Q2_cPopulatePanelTable)
            cnxn.commit()
            Q3_cPopulatePanelTable = "Update Panel SET Panel.Populated = 1 WHERE Panel.PanelID = %s" % (Cur_PanelID)
            cursor.execute(Q3_cPopulatePanelTable)
            cnxn.commit()
            Q_MCHigehestValue = "select cLastUsedValues.HighValue from cLastUsedValues"
            cursor.execute(Q_MCHigehestValue)
            HighestValue = cursor.fetchall()
            if Cur_PanelID > HighestValue:
                Q_Replace_MCHV = "Update cLastUsedValues SET cLastUsedValues.HighValue = %s" % (Cur_PanelID)
                cursor.execute(Q_Replace_MCHV)
                cnxn.commit()
            else:
                pass
        cursor.close()
        print "Done"

    def cPopulateCoreTable(self):
        print "Working..."
        print "Populate Core Table..."
        global HighestValue
        HighestValue = ()
        cursor = cnxn.cursor()
        Q1_cPopulateCoreTable = "select * from Core WHERE (Core.Populated is NULL)"
        cursor.execute(Q1_cPopulateCoreTable)
        Q1_result = cursor.fetchall()
        for CORE in Q1_result:
            if CORE.GRP == None:
                Cur_GRP = 'NULL'
            else:
                Cur_GRP = CORE.GRP
            if CORE.Ord == None:
                Cur_Ord = 'NULL'
            else:
                Cur_Ord = CORE.Ord
            if CORE.Core == None:
                Cur_Core = 'NULL'
            else:
                Cur_Core = "'" + str(CORE.Core) + "'"
            if CORE.Sleeve == None:
                Cur_Sleeve = 'NULL'
            else:
                Cur_Sleeve = "'" + str(CORE.Sleeve) + "'"
            if CORE.CoreTag1 == None:
                Cur_CoreTag1 = 'NULL'
            else:
                Cur_CoreTag1 = "'" + str(CORE.CoreTag1) + "'"
            if CORE.CoreTag2 == None:
                Cur_CoreTag2 = 'NULL'
            else:
                Cur_CoreTag2 = "'" + str(CORE.CoreTag2) + "'"
            if CORE.CoreTag3 == None:
                Cur_CoreTag3 = 'NULL'
            else:
                Cur_CoreTag3 = "'" + str(CORE.CoreTag3) + "'"
            if CORE.Network == None:
                Cur_Network = 'NULL'
            else:
                Cur_Network = "'" + str(CORE.Network) + "'"
            if CORE.Rev == None:
                Cur_Rev = 'NULL'
            else:
                Cur_Rev = "'" + str(CORE.Rev) + "'"
            if CORE.CableID == None:
                Cur_CableID = 'NULL'
            else:
                Cur_CableID = CORE.CableID
            if CORE.TstripID == None:
                Cur_TrayID = 'NULL'
            else:
                Cur_TrayID = CORE.TstripID
            if CORE.CoreID == None:
                Cur_CoreID = 'NULL'
            else:
                Cur_CoreID = CORE.CoreID
            Q2_cPopulateCoreTable = "INSERT INTO cCore (GRP,Ord,Core,Sleeve,CoreTag1,CoreTag2,CoreTag3,Network,Rev,CableID,TrayID,CoreID) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)" % (
            Cur_GRP, Cur_Ord, Cur_Core, Cur_Sleeve, Cur_CoreTag1, Cur_CoreTag2, Cur_CoreTag3, Cur_Network, Cur_Rev,
            Cur_CableID, Cur_TrayID, Cur_CoreID)
            cursor.execute(Q2_cPopulateCoreTable)
            cnxn.commit()
            Q3_cPopulateCoreTable = "Update Core SET Core.Populated = 1 WHERE Core.CoreID = %s" % (Cur_CoreID)
            cursor.execute(Q3_cPopulateCoreTable)
            cnxn.commit()
            Q_MCHigehestValue = "select cLastUsedValues.HighValue from cLastUsedValues"
            cursor.execute(Q_MCHigehestValue)
            HighestValue = cursor.fetchall()
            if Cur_CoreID > HighestValue:
                Q_Replace_MCHV = "Update cLastUsedValues SET cLastUsedValues.HighValue = " + "'" + Cur_CoreID + "'"
                cursor.execute(Q_Replace_MCHV)
                cnxn.commit()
            else:
                pass
        cursor.close()
        print "Done"

    def cUpdateDrawingSign(self):
        print "Working..."
        print "Batch update Revision Signatures ..."

    def cPopulateDetailCon1(self):
        print "Working..."
        print "Populate Detail Connection Table per Cable ..."
        global HighestValue
        HighestValue = ()
        cursor = cnxn.cursor()
        Q1_cPopulateDetailCon1 = "select Panel.TagNum as ScrPanelTagNum, DetailCon.ConID, DetailCon.TstripID as SrcTstripID, SRCTstrip.Display as SrcTstripTagNumDisplay, DetailCon.TermID as SrcTermID, SRCTerm.TNum as SrcTermTNum, SRCTstrip.TsCon as SrcTSCon, DetailCon.CoreID, Core.Core as CoreCol, Core.Sleeve as CoreSleeve, DetailCon.CableID, Cable.TagNum, DetailCon.DTstripID, DesTstrip.Display as DesTstripTagNumDisplay, DetailCon.DTermID, DesTerm.TNum as DesTermTNum, DesTstrip.TsCon as DesTSCon, (CASE WHEN SRCTstrip.TsCon = 'T' THEN SRCTstrip.AssocTstripID ELSE SRCTstrip.TstripID END) as UpdatedScrTstripID, (CASE WHEN SRCTstrip.TsCon = 'T' THEN DesTstrip.AssocTstripID ELSE DesTstrip.TstripID END) as UpdatedDesTstripID, SRCTstrip.Display as UpdatedScrTstripDisplay, DesTstrip.Display as UpdatedDesTstripDisplay, (CASE WHEN SRCTstrip.TsCon = 'T' THEN SRCTerm.TNum2 ELSE SRCTerm.TermID END) as UpdatedScrTermID, (CASE WHEN DesTstrip.TsCon = 'T' THEN DesTerm.TNum2 ELSE DesTerm.TermID END) as UpdatedDesTermID, SrcTerm.TNum as UpdatedScrTNum, DesTerm.TNum as UpdatedDesTNum from (((((((DetailCon inner join Term as SRCTerm on SRCTerm.TermID = DetailCon.TermID) inner join Term as DesTerm on DesTerm.TermID = DetailCon.DTermID) left join Tstrip as SRCTstrip on DetailCon.TstripID = SRCTstrip.TstripID) left join Tstrip as DesTstrip on DetailCon.DTstripID = DesTstrip.TstripID) left join Panel on SrcTstrip.AssocGID = Panel.PanelID) inner join Core on Core.CoreID = DetailCon.CoreID) inner join Cable on Cable.CableID = DetailCon.CableID) where (DetailCon.CableID > 0 AND DetailCon.Populated is NULL)"
        cursor.execute(Q1_cPopulateDetailCon1)
        Q1_result = cursor.fetchall()
        for DC in Q1_result:
            if DC.ConID == None:
                Cur_ConID = 'NULL'
            else:
                Cur_ConID = DC.ConID
            if DC.SrcTstripID == None:
                Cur_SrcTstripID = 'NULL'
            else:
                Cur_SrcTstripID = DC.SrcTstripID
            if DC.SrcTstripTagNumDisplay == None:
                Cur_SrcTstripTagNumDisplay = 'NULL'
            else:
                Cur_SrcTstripTagNumDisplay = "'" + str(DC.SrcTstripTagNumDisplay) + "'"
            if DC.SrcTermID == None:
                Cur_SrcTermID = 'NULL'
            else:
                Cur_SrcTermID = DC.SrcTermID
            if DC.SrcTermTNum == None:
                Cur_SrcTermTNum = 'NULL'
            else:
                Cur_SrcTermTNum = "'" + str(DC.SrcTermTNum) + "'"
            if DC.CoreID == None:
                Cur_CoreID = 'NULL'
            else:
                Cur_CoreID = DC.CoreID
            if DC.CoreCol == None:
                Cur_CoreCol = 'NULL'
            else:
                Cur_CoreCol = "'" + str(DC.CoreCol) + "'"
            if DC.CoreSleeve == None:
                Cur_CoreSleeve = 'NULL'
            else:
                Cur_CoreSleeve = "'" + str(DC.CoreSleeve) + "'"
            if DC.CableID == None:
                Cur_CableID = 'NULL'
            else:
                Cur_CableID = DC.CableID
            if DC.TagNum == None:
                Cur_CableTagNum = 'NULL'
            else:
                Cur_CableTagNum = "'" + str(DC.TagNum) + "'"
            if DC.DTstripID == None:
                Cur_DTstripID = 'NULL'
            else:
                Cur_DTstripID = DC.DTstripID
            if DC.DesTstripTagNumDisplay == None:
                Cur_DesTstripTagNumDisplay = 'NULL'
            else:
                Cur_DesTstripTagNumDisplay = "'" + str(DC.DesTstripTagNumDisplay) + "'"
            if DC.DTermID == None:
                Cur_DesTermID = 'NULL'
            else:
                Cur_DesTermID = DC.DTermID
            if DC.DesTermTNum == None:
                Cur_DesTermTNum = 'NULL'
            else:
                Cur_DesTermTNum = "'" + str(DC.DesTermTNum) + "'"
            Q2_cPopulateDetailConTable = "INSERT INTO cDetailCon (ConID, SrcTstripID, SrcTstripTagNum, SrcTermID, SrcTermTNum, CoreID, SrcCoreCol, SrcSleeveCol, CableID, CableTagNum, DTstripID, DesTstripTagNum, DTermID, DesTermTNum) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)" % (
            Cur_ConID, Cur_SrcTstripID, Cur_SrcTstripTagNumDisplay, Cur_SrcTermID, Cur_SrcTermTNum, Cur_CoreID,
            Cur_CoreCol, Cur_CoreSleeve, Cur_CableID, Cur_CableTagNum, Cur_DTstripID, Cur_DesTstripTagNumDisplay,
            Cur_DesTermID, Cur_DesTermTNum)
            cursor.execute(Q2_cPopulateDetailConTable)
            cnxn.commit()
            Q3_cPopulateDetailConTable = "Update DetailCon SET DetailCon.Populated = 1 WHERE DetailCon.TermID = %s" % (
            Cur_SrcTermID)
            cursor.execute(Q3_cPopulateDetailConTable)
            cnxn.commit()
            Q_MCHigehestValue = "select cLastUsedValues.HighValue from cLastUsedValues"
            cursor.execute(Q_MCHigehestValue)
            HighestValue = cursor.fetchall()
            if Cur_ConID > HighestValue:
                Q_Replace_MCHV = "Update cLastUsedValues SET cLastUsedValues.HighValue = %s" % (Cur_ConID)
                cursor.execute(Q_Replace_MCHV)
                cnxn.commit()
            else:
                pass
        cursor.close()
        print "Done"

    def cPopulateDrawingNumbers_Patch(self):
        print "Working..."
        print "Populate Detail Connection Table per Cable ..."

        # Get highest unique ID
        global newhighvalue
        global HighestValue
        global NullValue
        NullValue = 'NULL'
        HighestValue = ()
        # newhighvalue = self.cGetHighUniqueID()

        projiniFileName = 'projini.txt'
        projiniFileNameOld = 'projini_old.txt'
        now = time.time()
        userName = getpass.getuser()
        currentDirectory = os.getcwd()
        currentDateTime = time.time()
        pnumber = os.getpid()
        fullFilePathName = (currentDirectory + '/' + projiniFileName)
        fullFilePathNameOld = (currentDirectory + '/' + projiniFileNameOld)

        groupY = (
        514, 499, 484, 469, 454, 439, 424, 409, 394, 379, 364, 349, 334, 319, 304, 289, 274, 259, 244, 229, 214, 199,
        184, 175, 169, 154)
        groupX = {'11': 309, '21': 309, '22': 469, '31': 229, '32': 389, '33': 549, '41': 149, '42': 309, '43': 469,
                  '44': 629, '51': 69, '52': 229, '53': 389, '54': 549, '55': 709, '61': 69, '62': 229, '63': 389,
                  '64': 549, '65': 709, '66': 869, '71': 69, '72': 229, '73': 389, '74': 549, '75': 709, '76': 869,
                  '77': 1029, '81': 69, '82': 229, '83': 389, '84': 549, '85': 709, '86': 869, '87': 1029, '88': 1189}

        cursor = cnxn.cursor()

        Cur_Area = '07GB'

        # Get list of all panels:
        Qry1 = "select cPanel.* from cPanel"
        cursor.execute(Qry1)
        Qry1_R = cursor.fetchall()
        for Pan in Qry1_R:
            newhighvalue = (self.cGetHighUniqueID() + 1)
            Cur_PanelID = Pan.PanelID
            Cur_DwgNr = Pan.Notes
            Cur_Pan_TagNum = Pan.TagNum
            Cur_Application = 'FDES'
            Cur_DType = 'DWG'
            Cur_Template = 'Patch_FO.dwg'
            Cur_PrintTable = 'PANEL'
            print "\nCurrent Panel: %s" % (Cur_PanelID)
            print "Current Drawing: %s" % (Cur_DwgNr)
            Qry2 = "select CC.CableID, CC.CableTagSource, CC.CableTagDestination, CP.TagNum, CP.Notes from ((cCable as CC left join CableCon as CCC on CC.CableID = CCC.CableID) left join cPanel as CP on CP.PanelID = CCC.SrcID) where CP.PanelID = %s" % (
            Cur_PanelID)
            cursor.execute(Qry2)
            Qry2_R = cursor.fetchall()
            print "Amount of Outgoing Cables for this Panel: %s" % (len(Qry2_R))
            Qry3 = "select CC.CableID, CC.CableTagSource, CC.CableTagDestination, CP.TagNum from ((cCable as CC left join CableCon on CC.CableID = CableCon.CableID) left join Panel as CP on CP.PanelID = CableCon.DesID) where CP.PanelID = %s" % (
            Cur_PanelID)
            cursor.execute(Qry3)
            Qry3_R = cursor.fetchall()
            print "Amount of Incoming Cables for this Panel: %s" % (len(Qry3_R))
            Qry4 = "select cTray.* from cTray where cTray.AssocID = %s ORDER by cTray.Ord" % (Cur_PanelID)
            cursor.execute(Qry4)
            Qry4_R = cursor.fetchall()

            print "Amount of Splice Trays for this Panel: %s" % (len(Qry4_R))

            # Create a new drawing for Patching details:
            print "New Patching Drawing: 120108-DWG-07GB-108-%s-0%s" % ((Cur_DwgNr), (len(Qry2_R) + 1))
            NewPatchDwgNr = "120108-DWG-07GB-108-%s-0%s" % ((Cur_DwgNr), (len(Qry2_R) + 1))
            NewShtNr = "0%s" % (len(Qry2_R) + 1)
            print "---"
            global printRange
            printRange = []

            Cur_Dwg_Desc = "Data Communications - Panel: %s - Fibre Patch Diagram" % (Cur_Pan_TagNum)

            for tray in Qry4_R:
                Cur_TrayID = tray.TrayID
                Cur_TrayDisplay = tray.Display
                Cur_TrayOrd = tray.Ord
                Xcoordinate = '%s%s' % ((len(Qry4_R)), (Qry4_R.index(tray) + 1))
                Cur_DocID = tray.DocumentID
                printRange.append(Cur_TrayID)
                # print "Current TrayID: %s" %(Cur_TrayID)
                # print "Current Tray Display: %s" %(Cur_TrayDisplay)
                # print "Current Tray Order: %s" %(Cur_TrayOrd)
                # print "X-coordinate: %s" %(Xcoordinate)
                # print groupX[Xcoordinate]
                Qry5 = "UPDATE cTray SET cTray.PatchX = %s WHERE cTray.TrayID = %s" % (
                (groupX[Xcoordinate]), (Cur_TrayID))
                print Qry5
                cursor.execute(Qry5)
                cnxn.commit()
                Qry6 = "UPDATE cTray SET cTray.Notes = '%s' WHERE cTray.TrayID = %s" % ((NewPatchDwgNr), (Cur_TrayID))
                print Qry6
                cursor.execute(Qry6)
                cnxn.commit()
                Qry10 = "UPDATE cTray SET cTray.DocumentID = %s WHERE cTray.TrayID = '%s'" % (newhighvalue, Cur_TrayID)
                print Qry10
                cursor.execute(Qry10)
                cnxn.commit()

            if Cur_DocID == None:
                Qry7 = "INSERT INTO Document (TagNum, Description, FileName ,DwgName, SheetNo, DType, Area, Application, Template, PrintTable,PrintID, DocumentID) VALUES('%s', '%s', '%s', '%s', '%0s', '%s', '%s', '%s', '%s', '%s', %s, %s)" % (
                Cur_Pan_TagNum, Cur_Dwg_Desc, NewPatchDwgNr, Cur_DwgNr, NewShtNr, Cur_DType, Cur_Area, Cur_Application,
                Cur_Template, Cur_PrintTable, Cur_PanelID, newhighvalue)

                print Qry7
                cursor.execute(Qry7)
                cnxn.commit()

                Qry9 = "Update zLastUsedID SET LastID = %s FROM zLastUsedID" % (newhighvalue)
                cursor.execute(Qry9)
                cnxn.commit()
            else:
                pass

            QtyTrays = len(Qry4_R)
            rangetoprint = printRange[0]
            Qry11 = "UPDATE Document SET Document.PrintRange = '%s' WHERE Document.DocumentID = '%s'" % (
            rangetoprint, newhighvalue)
            print Qry11
            cursor.execute(Qry11)
            cnxn.commit()

            if QtyTrays > 1:
                for Q in range(1, QtyTrays):
                    Qry12 = "select Document.* from Document WHERE Document.DocumentID = %s" % (newhighvalue)
                    cursor.execute(Qry12)
                    Qry12_R = cursor.fetchone()
                    Cur_printRange = Qry12_R.PrintRange
                    New_printRange = "%s,%s" % (Cur_printRange, printRange[Q])
                    Qry13 = "UPDATE Document SET Document.PrintRange = '%s' WHERE Document.DocumentID = '%s'" % (
                    New_printRange, newhighvalue)
                    print Qry13
                    cursor.execute(Qry13)
                    cnxn.commit()

            else:
                pass

        cursor.close()
        print "Done"

    def cGetHighUniqueID(self):
        global HighestValue
        global Q1_result
        print "Working..."
        print "Getting the highest unique ID..."
        cursor = cnxn.cursor()
        Q1 = "SELECT zLastUsedID.* FROM zLastUsedID"
        cursor.execute(Q1)
        # Q1_result = cursor.fetchall()
        # print Q1_result
        HighestValue = cursor.fetchone()
        return (int(HighestValue[0]) + 1)
        # Q1_cGetHighUniqueID = "select max(maxID) FROM (select max(ConID) AS maxID from cDetailCon union select max(CableID) AS maxID from cCable union select max(ConID) AS maxID from cCableCon union select max(CoreID) AS maxID from cCore union select max(PanelID) AS maxID from cPanel union select max(TrayID) AS maxID from cTray union select max(DocumentID) AS maxID from Document union select max(MidCouplerID) AS maxID from cMidCoupler union select max(LoopID) AS maxID from Loop union select max(IIndexID) AS maxID from IIndex) as SubQuery"
        # cursor.execute(Q1_cGetHighUniqueID)
        # Q1_result = cursor.fetchone()
        # Q_MCHigehestValue = "select cLastUsedValues.HighValue from cLastUsedValues"
        # cursor.execute(Q_MCHigehestValue)
        # HighestValue = cursor.fetchone()
        # if Q1_result >= HighestValue:
        #	Q_Replace_MCHV = "Update cLastUsedValues SET cLastUsedValues.HighValue = %s" %(Q1_result[0])
        #	cursor.execute(Q_Replace_MCHV)
        #	cnxn.commit()
        #	return (Q1_result[0])
        # else:
        #	return (HighestValue[0])

        print "Done"
        cursor.close()

    def cCreatePatchDocuments(self):
        newhighvalue = self.cGetHighUniqueID()
        print newhighvalue

    def PopulateIPaddressDetails(self):
        print "Working..."
        print "Populate IP address Details"
        cursor = cnxn.cursor()
        Q1_PopulateIPaddressDetails = "UPDATE IIndex SET IIndex.IP_a = SUBSTRING(IIndex.TagNum,1,3) FROM IIndex"
        cursor.execute(Q1_PopulateIPaddressDetails)
        cnxn.commit()
        Q2_PopulateIPaddressDetails = "UPDATE IIndex SET IIndex.IP_b = SUBSTRING(IIndex.TagNum,5,3) FROM IIndex"
        cursor.execute(Q2_PopulateIPaddressDetails)
        cnxn.commit()
        Q3_PopulateIPaddressDetails = "UPDATE IIndex SET IIndex.IP_c = SUBSTRING(IIndex.TagNum,9,3) FROM IIndex"
        cursor.execute(Q3_PopulateIPaddressDetails)
        cnxn.commit()
        Q4_PopulateIPaddressDetails = "UPDATE IIndex SET IIndex.IP_d = SUBSTRING(IIndex.TagNum,13,3) FROM IIndex"
        cursor.execute(Q4_PopulateIPaddressDetails)
        cnxn.commit()
        cursor.close()
        print "Done"

    def PopulateNetworkTable(self):
        print "Working..."
        print "Populate Network Table"
        cursor = cnxn.cursor()
        Q1_PopulateNetworkTable = "select IIndex.Area FROM IIndex WHERE IIndex.User1 is NULL Group By IIndex.Area"
        cursor.execute(Q1_PopulateNetworkTable)
        Networks = cursor.fetchall()
        for NW in Networks:
            Cur_NW = str(NW.Area)
            Q2_PopulateNetworkMasters = "INSERT TOP(1) INTO Network select IIndex.Area, IIndex.TagNum, IIndex.IP_a, IIndex.IP_b, IIndex.IP_c, IIndex.IIndexID, IIndex.PanelID From IIndex Left Join Panel ON Panel.PanelID = IIndex.PanelID WHERE IIndex.Area = " + "'" + str(
                Cur_NW) + "'" + " ORDER BY IIndex.IIndexID"
            cursor.execute(Q2_PopulateNetworkMasters)
            cnxn.commit()
            Q3_PopulateNetworkTable = "Update IIndex set IIndex.User1 = IIndex.Area WHERE IIndex.Area = " + "'" + str(
                Cur_NW) + "'"
            cursor.execute(Q3_PopulateNetworkTable)
            cnxn.commit()

            '''
            This needs to be done:

            select IIndex.Area
            FROM IIndex
            Group By IIndex.Area

            select TOP(1)(IIndex.TagNum), IIndex.IP_a, IIndex.IP_b, IIndex.IP_c, IIndex.IP_d,
            IIndex.Area, IIndex.PanelID, IIndex.IIndexID
            From IIndex
            Inner Join Panel ON Panel.PanelID = IIndex.PanelID
            WHERE IIndex.Area = 'R23'
            ORDER BY IIndex.IIndexID

            '''

        cursor.close()
        print "Done"

    def AssociateCableswithDrawingsTemplates(self):
        cursor = cnxn.cursor()

        def populatePaneldwgdetails(self):
            q1_panels = "SELECT Panel.TagNum, Panel.OtherDocs, Panel.PanelID FROM Panel ORDER BY Panel.PanelID"
            cursor.execute(q1_panels)
            q1_panels_result = cursor.fetchall()
            count_q1 = 800
            for q1 in q1_panels_result:
                q1_cur_PanelID = q1.PanelID
                count_q1 += 1
                q1_updatequery = "UPDATE Panel SET Panel.Notes =" + "'0" + str(
                    count_q1) + "' WHERE Panel.PanelID = " + "'" + str(q1_cur_PanelID) + "'"
                cursor.execute(q1_updatequery)
                cnxn.commit()
                q2_cable = "SELECT CableCon.CableID, CableCon.SrcID, Panel.PanelID, Panel.Notes, substring(Cable.CableSTD,5,2) AS CoreCount FROM ((CableCon inner join Panel on Panel.PanelID = CableCon.SrcID) inner join Cable ON Cable.CableID = CableCon.CableID) WHERE CableCon.CableID <> '' AND Panel.PanelID = '" + str(
                    q1_cur_PanelID) + "'"
                cursor.execute(q2_cable)
                q2_cable_result = cursor.fetchall()
                count_q2 = 0
                for q2 in q2_cable_result:
                    q2_cur_corecount = q2.CoreCount
                    q2_cur_CableID = q2.CableID
                    q2_cur_PanelID = q2.PanelID
                    q2_cur_PanelDwgnr = q2.Notes
                    count_q2 += 1
                    if count_q2 < 10:
                        countnumber = '0' + str(count_q2)
                    else:
                        countnumber = count_q2
                    q2_updatequery = "UPDATE Cable SET Cable.Notes =" + "'" + str(q2_cur_PanelDwgnr) + "-" + str(
                        countnumber) + "-" + str(q2_cur_PanelID) + "-" + str(q2_cur_corecount) + "-" + str(
                        q2_cur_CableID) + "' FROM Cable WHERE Cable.CableID = " + "'" + str(q2_cur_CableID) + "'"
                    cursor.execute(q2_updatequery)
                    cnxn.commit()

        def populateTemplateCableCores(self):
            q3_cable = "SELECT Cable.CableID, Cable.Notes, Core.CoreID, Core.GRP, Core.Ord, DetailCon.TermID, Term.TNum, Term.TstripID, Tstrip.Display, Tstrip.TsCon FROM (((Cable inner join Core on Core.CableID = Cable.CableID) inner join DetailCon on DetailCon.CoreID = Core.CoreID) inner join Term on DetailCon.TermID = Term.TermID) inner join Tstrip ON Tstrip.TstripID = Term.TstripID ORDER BY Cable.CableID, Core.Ord"
            cursor.execute(q3_cable)
            q3_cable_result = cursor.fetchall()
            for q3 in q3_cable_result:
                q3_curr_termnum = q3.TNum
                q3_curr_Display = q3.Display
                q3_curr_ord = q3.Ord
                q3_curr_tsID = q3.TstripID
                q3_curr_CableID = q3.CableID
                q3_curr_CoreID = q3.CoreID
                q3_curr_Notes = q3.Notes

                q3_dwgnr = int(q3_curr_Notes.split('-')[0])
                q3_shtnr = int(q3_curr_Notes.split('-')[1])
                q3_panelnr = int(q3_curr_Notes.split('-')[2])
                q3_cur_cable_amtCores = int(q3_curr_Notes.split('-')[3])

                if q3_curr_ord == 1:
                    CableCoreTerm1 = q3.TNum
                    partA_tsID = q3.TstripID
                if q3_curr_ord == ((q3_cur_cable_amtCores) / 2) + 1:
                    partB_tsID = q3.TstripID
                    if partA_tsID == partB_tsID:
                        q4a = "UPDATE Cable SET Cable.Notes= '0" + str(q3_dwgnr) + "-0" + str(q3_shtnr) + "-" + str(
                            q3_panelnr) + "-" + str(q3_cur_cable_amtCores) + "-" + str(
                            CableCoreTerm1) + "-FTFT" + "-" + str(q3_curr_CableID) + "-" + str(
                            partA_tsID) + "' WHERE Cable.CableID = " + "'" + str(q3_curr_CableID) + "'"
                        cursor.execute(q4a)
                        cnxn.commit()
                        q4aa = "UPDATE Cable SET Cable.CableTemplate= 'Active' WHERE Cable.CableID = " + "'" + str(
                            q3_curr_CableID) + "'"
                        cursor.execute(q4aa)
                        cnxn.commit()
                    else:
                        q4b = "UPDATE Cable SET Cable.Notes= '0" + str(q3_dwgnr) + "-0" + str(q3_shtnr) + "-" + str(
                            q3_panelnr) + "-" + str(q3_cur_cable_amtCores) + "-" + str(
                            CableCoreTerm1) + "-FTTF" + "-" + str(q3_curr_CableID) + "-" + str(partA_tsID) + "," + str(
                            partB_tsID) + "' WHERE Cable.CableID = " + "'" + str(q3_curr_CableID) + "'"
                        cursor.execute(q4b)
                        cnxn.commit()
                        q4bb = "UPDATE Cable SET Cable.CableTemplate= 'Active' WHERE Cable.CableID = " + "'" + str(
                            q3_curr_CableID) + "'"
                        cursor.execute(q4bb)
                        cnxn.commit()
                else:
                    pass

        def populateDrawingdetails(self):
            q5 = "SELECT Cable.CableID, Cable.Notes, Cable.CableTag1, Cable.CableTag2, Document.DocumentID, Document.PrintID FROM Cable inner join Document ON Document.PrintID = Cable.CableID WHERE Cable.CableTemplate= 'Active'"
            cursor.execute(q5)
            q5_result = cursor.fetchall()

            for q5 in q5_result:
                q5_cabTag1 = q5.CableTag1
                q5_cabTag2 = q5.CableTag2
                q5_notes = q5.Notes
                q5_docID = q5.DocumentID
                q5_cabID = q5.CableID
                q5_dwgnr = (q5_notes.split('-')[0])
                q5_shtnr = (q5_notes.split('-')[1])
                q5_panelID = (q5_notes.split('-')[2])
                q5_cores = (q5_notes.split('-')[3])
                q5_firstcore = (q5_notes.split('-')[4])
                q5_assembly = (q5_notes.split('-')[5])
                q5_cableID = (q5_notes.split('-')[6])
                q5_printrange = (q5_notes.split('-')[7])

                q6a = "UPDATE Document SET Document.FileName = '120108-DWG-07GB-108-" + str(q5_dwgnr) + "-" + str(
                    q5_shtnr) + "' WHERE Document.DocumentID = " + "'" + str(q5_docID) + "'"
                cursor.execute(q6a)
                cnxn.commit()
                q6b = "UPDATE Document SET Document.DType = 'DWG' WHERE Document.DocumentID = " + "'" + str(
                    q5_docID) + "'"
                cursor.execute(q6b)
                cnxn.commit()
                q6c = "UPDATE Document SET Document.PrintTable = 'PANEL' WHERE Document.DocumentID = " + "'" + str(
                    q5_docID) + "'"
                cursor.execute(q6c)
                cnxn.commit()
                q6d = "UPDATE Document SET Document.PrintID = '" + str(
                    q5_panelID) + "' WHERE Document.DocumentID = " + "'" + str(q5_docID) + "'"
                cursor.execute(q6d)
                cnxn.commit()
                q6e = "UPDATE Document SET Document.PrintRange = " + "'" + str(
                    q5_printrange) + "' WHERE Document.DocumentID = " + "'" + str(q5_docID) + "'"
                cursor.execute(q6e)
                cnxn.commit()
                q6f = "UPDATE Document SET Document.MainTemplate = " + "'" + str(q5_cores) + "-" + str(
                    q5_firstcore) + "-" + str(q5_assembly) + "' WHERE Document.DocumentID = " + "'" + str(
                    q5_docID) + "'"
                cursor.execute(q6f)
                cnxn.commit()
                q6g = "UPDATE Document SET Document.SheetNo = '01' WHERE Document.DocumentID = " + "'" + str(
                    q5_docID) + "'"
                cursor.execute(q6g)
                cnxn.commit()
                q6h = "UPDATE Document SET Document.DwgName = '" + str(
                    q5_dwgnr) + "' WHERE Document.DocumentID = " + "'" + str(q5_docID) + "'"
                cursor.execute(q6h)
                cnxn.commit()
                q6i = "UPDATE Document SET Document.Area = '07GB' WHERE Document.DocumentID = " + "'" + str(
                    q5_docID) + "'"
                cursor.execute(q6i)
                cnxn.commit()
                q6j = "UPDATE Document SET Document.Description = 'Fibre Cable Splicing Detail: " + str(
                    q5_cabTag1) + "/" + str(q5_cabTag2) + "' WHERE Document.DocumentID = " + "'" + str(q5_docID) + "'"
                cursor.execute(q6j)
                cnxn.commit()

        populatePaneldwgdetails(self)
        populateTemplateCableCores(self)
        populateDrawingdetails(self)
        print " done"

        cursor.close()

        '''
        for q1 in q_detailcon_result:
            Cur_CableID = q1.CableID
            #Cur_CableCon_SrcID = q1.SrcID
            Cur_Panel_OtherDocs = q1.OtherDocs
            Cur_TstripID = q1.TstripID
            Cur_CoreID = q1.CoreID
            Cur_Ord = q1.Ord
            Cur_PanelID = q1.PanelID

            if Cur_Ord == 1:
                print "-----"
                print Cur_CableID
                print Cur_Panel_OtherDocs
                print Cur_TstripID
                print Cur_CoreID
                print Cur_Ord
                print Cur_PanelID
            else:
                pass




            Q2_ = "SELECT Core.CoreID, Core.Ord FROM Core WHERE Core.CableID = " + "'" + str(Cur_CableID) + "'"
            cursor.execute(Q2_)
            Q2_result = cursor.fetchall()
            sumQ2 = len(Q2_result)

            for q2 in Q2_result:
                Cur_CoreID = (q2.CoreID)
                Cur_CoreOrd = (q2.Ord)

                if Cur_CoreOrd == 1:
                    Q3a_ = "SELECT DetailCon.ConID, DetailCon.CoreID, DetailCon.TstripID, DetailCon.CableID, Tstrip.AssocGID FROM (DetailCon LEFT JOIN Tstrip ON Tstrip.TstripID = DetailCon.TstripID) WHERE DetailCon.CoreID = " + "'" + str(Cur_CoreID) +"'"
                    cursor.execute(Q3a_)
                    Q3a_result = cursor.fetchall()
                    for q3a in Q3a_result:
                        Curr_ConIDa = (q3a.ConID)
                        Curr_CoreIDa = (q3a.CoreID)
                        Curr_TstripIDa = (q3a.TstripID)
                        Curr_CableIDa = (q3a.CableID)
                        Curr_AssocGIDa = (q3a.AssocGID)
                        print Cur_CoreOrd, Curr_ConIDa, Curr_CoreIDa, Curr_TstripIDa, Curr_CableIDa, Curr_AssocGIDa

                if Cur_CoreOrd == (sumQ2/2):

                    Q3b_ = "SELECT DetailCon.ConID, DetailCon.CoreID, DetailCon.TstripID, DetailCon.CableID, Tstrip.AssocGID FROM (DetailCon LEFT JOIN Tstrip ON Tstrip.TstripID = DetailCon.TstripID) WHERE DetailCon.CoreID = " + "'" + str(Cur_CoreID) +"'"
                    cursor.execute(Q3b_)
                    Q3b_result = cursor.fetchall()

                    for q3b in Q3b_result:
                        Curr_ConIDb = (q3b.ConID)
                        Curr_CoreIDb = (q3b.CoreID)
                        Curr_TstripIDb = (q3b.TstripID)
                        Curr_CableIDb = (q3b.CableID)
                        Curr_AssocGIDb = (q3b.AssocGID)
                        print Cur_CoreOrd, Curr_ConIDb, Curr_CoreIDb, Curr_TstripIDb, Curr_CableIDb, Curr_AssocGIDb
                else:
                    pass





                    #C_curr_tsID = q3.TstripID
                    #C_curr_assocGID = q3.AssocGID
                    #if Cur_CoreOrd == 1:
                #		print Cur_CoreOrd, C_curr_tsID, C_curr_assocGID
                #	if Cur_CoreOrd == (sumQ2/2):
                #		print Cur_CoreOrd, C_curr_tsID, C_curr_assocGID
                #	else:
                #		pass


                    #if Q3_aresult == 'Null':
                    #	pass
                    #else:
                        #pass
                    #	print CoreID
                    #	print TstripID
                    #	print AssocGID

                    if C_Ord == int(sumQ2/2):
                        Q3_b = "SELECT DetailCon.CoreID, DetailCon.TstripID FROM DetailCon WHERE (DetailCon.CoreID = " + "'" + str(C_CoreID) + "' AND DetailCon.TstripID <> " + "'" + "'"+")"
                        #print Q3_b
                        cursor.execute(Q3_b)
                        Q3_bresult = cursor.fetchone()
                        if Q3_bresult == 'Null':
                            pass
                            #print "null"
                        else:
                            pass
                        #print Q3_bresult
                        #partB = TstripID

                        if Q3_aresult == Q3_bresult:
                            CableTemplate = 'FO_FFTT' + str(sumQ2)+'.dwg'
                        else:
                            CableTemplate = 'FO_FTTF' + str(sumQ2)+'.dwg'
                        #print CableTemplate

                    Q_UpdateCableTemplate = "Update Cable SET Cable.CableTemplate = " +"'" + str(CableTemplate) +"'"+" from Cable WHERE CableID = '" + str(C_CableID) + "'"
                    print Q_UpdateCableTemplate
                    cursor.execute(Q_UpdateCableTemplate)
                    cnxn.commit()

                else:
                    pass
        '''

    def AllocateDwgNrs(self):
        cursor = cnxn.cursor()

        queryAllocateDwgNrs1 = "Select Cable.CableID, Panel.OtherDocs, Panel.PanelID, CableCon.CableID, CableCon.SrcID from ((Cable inner join CableCon on Cable.CableID = CableCon.CableID) inner join Panel on Panel.PanelID = CableCon.SrcID) order by Panel.PanelID, Cable.CableID"
        cursor.execute(queryAllocateDwgNrs1)
        queryAllocate1 = cursor.fetchall()

        # print queryAllocate1
        for qryAllocate1 in queryAllocate1:
            Q1_Cur_PanelID = str(qryAllocate1.PanelID)
            queryAllocateDwgNrs2 = "Select Cable.CableID, Panel.OtherDocs, Panel.PanelID, CableCon.CableID, CableCon.SrcID from ((Cable inner join CableCon on Cable.CableID = CableCon.CableID) inner join Panel on Panel.PanelID = CableCon.SrcID) where Panel.PanelID = " + "'" + Q1_Cur_PanelID + "'" + " order by Panel.PanelID, Cable.CableID"
            cursor.execute(queryAllocateDwgNrs2)
            queryAllocate2 = cursor.fetchall()

            for qryAllocate2 in queryAllocate2:
                Q2_Cur_PanelOtherDocs = str(qryAllocate2.OtherDocs)
                Q2_Cur_CableID = str(qryAllocate2.CableID)
                Q2_series = (queryAllocate2.index(qryAllocate2)) + 1
                queryAllocateDwgNrs3a = "Update Cable SET Cable.OtherDocs = '120108-DGW120108-DWG-07GB-108-" + str(
                    Q2_Cur_PanelOtherDocs) + "-0" + str(
                    Q2_series) + "' from Cable WHERE CableID = '" + Q2_Cur_CableID + "'"
                cursor.execute(queryAllocateDwgNrs3a)
                cnxn.commit()
                queryAllocateDwgNrs3b = "Update Document SET Document.FileName = '120108-DGW120108-DWG-07GB-108-" + str(
                    Q2_Cur_PanelOtherDocs) + "-0" + str(
                    Q2_series) + "' from (Cable inner join Document ON Document.PrintID = Cable.CableID) WHERE CableID = '" + Q2_Cur_CableID + "'"
                cursor.execute(queryAllocateDwgNrs3b)
                cnxn.commit()
                print queryAllocateDwgNrs3a
                print queryAllocateDwgNrs3b

        # _C_CableSRD = C.CableSTD
        #	_C_Area = C.Area
        #	_C_Func = C.Func
        #	_C_Num = C.Num
        #	_C_Suffix = C.Suffix
        #	_C_OtherDocs = C.OtherDocs
        #	_C_EstLength = C.EstLength
        #
        #	_CableDict_+ str(C) = {}
        #	'_CableDict_'+ str(C))['index'] = C
        #	('_CableDict_'+ str(C))['Dict_nr'] = ('_CableDict_'+ str(C))
        #	('_CableDict_'+ str(C))['CableID'] = CableID
        #	print ('_CableDict_'+ str(C))
        cursor.close()

    def UpdateTableStructure(self):
        # Update Documents table & Revision table with custom fields
        cursor = cnxn.cursor()

        Q1_UpdateTable = "ALTER TABLE Document ADD [ToUpdate] varchar(12),[AreaDesc] varchar(50),[ProjectEngineer]varchar(10),[DocDescription] varchar(50),[PackageSOWNr] varchar(50),[DocCheckedBy] varchar(10),[DocCheckedBy_Date] varchar(10),[DocApprovedBy] varchar(10),[DocApprovedBy_Date] varchar(10),[DocDiscSL] varchar(10),[DocDiscSL_Date] varchar(10),[DocProjSL] varchar(10),[DocProjSL_Date] varchar(10),[DocCILeadEng] varchar(10),[DocCILeadEng_Date] varchar(10),[DocCIEng] varchar(10),[DocCIEng_Date] varchar(10),[DocPEM] varchar(10),[DocPEM_Date] varchar(10),[DocPM] varchar(10),[DocPM_Date] varchar(10),[DocC_PM] varchar(10),[DocC_PM_Date] varchar(10),[DocC_PEM] varchar(10),[DocC_PEM_Date] varchar(10),[DocC_CIEng] varchar(10),[DocC_CIEng_Date] varchar(10),[DocC_DesEng] varchar(10),[DocC_DesEng_Date] varchar(10),[PREngName] varchar(10),[PREngNr] varchar(10),[PREngRevSign] varchar(10)"
        cursor.execute(Q1_UpdateTable)
        cnxn.commit()
        Q2_UpdateTable = "ALTER TABLE Revision ADD CheckedBy_Date varchar(8), ApprovedBy_Date varchar(8), DiscSL varchar(4), DiscSL_Date varchar(8), ProjSL varchar(4), ProjSL_Date varchar(8), CILeadEng varchar(4), CILeadEng_Date varchar(8), CIEng varchar(4), CIEng_Date varchar(8), PEM varchar(4), PEM_Date varchar(8), PM varchar(4),PM_Date varchar(8), C_PM varchar(4), C_PM_Date varchar(8), C_PEM varchar(4), C_PEM_Date varchar(8), C_CIEng varchar(4), C_CIEng_Date varchar(8), C_DesEng varchar(4), C_DesEng_Date varchar(8)"
        cursor.execute(Q2_UpdateTable)
        cnxn.commit()
        Q3_UpdateTable = "ALTER TABLE Cable ADD CableTemplate varchar(20)"
        cursor.execute(Q3_UpdateTable)
        cnxn.commit()
        Q4_UpdateTable = "ALTER TABLE IIndex ADD IP_a varchar(5)"
        cursor.execute(Q4_UpdateTable)
        cnxn.commit()
        Q5_UpdateTable = "ALTER TABLE IIndex ADD IP_b varchar(5)"
        cursor.execute(Q5_UpdateTable)
        cnxn.commit()
        Q6_UpdateTable = "ALTER TABLE IIndex ADD IP_c varchar(5)"
        cursor.execute(Q6_UpdateTable)
        cnxn.commit()
        Q7_UpdateTable = "ALTER TABLE IIndex ADD IP_d varchar(5)"
        cursor.execute(Q7_UpdateTable)
        cnxn.commit()
        Q8_CreateTable = "Create Table Network (NetworkID int, NW_TagNum varchar(10), NW_IP_a varchar(3), NW_IP_b varchar(3), NW_IP_c varchar(3), NW_main_ID varchar(10), NW_startPanelID varchar(10))"
        cursor.execute(Q8_CreateTable)
        cnxn.commit()
        Q9_CreateTable = "Create Table NetworkRoute (NWRouteSeqID int, Network_ID varchar(10), NW_SourcePanelID varchar(10), NW_DesPanelID varchar(10))"
        cursor.execute(Q9_CreateTable)
        cnxn.commit()
        Q10_CreateTable = "Create Table NetworkConnectedPanels (ConnectionID int, NW_PanelID varchar(10), NW_ConnectedPanelID varchar(10))"
        cursor.execute(Q10_CreateTable)
        cnxn.commit()

        cursor.close()

    def ReadTableStructure(self):
        curs = cnxn.cursor()
        Tbls = []
        for row in cursor.tables():
            print "====================="
            # print row.table_name
            print (row.table_schem + '.' + row.table_name)
            for Column in row:
                print Column

                # print (row.table_schem + '.' + row.table_name)

        # for row in curs.tables(tableType = 'TABLE').fetchall():
        #	Tbls.append(row.table_schem + '.' + row.table_name)
        # print Tbls

        curs.close()

    def ImportDwg(self):
        from dxfgrabber.tags import Tags
        from dxfgrabber.blockssection import BlocksSection

        dxf = dxfgrabber.readfile(
            "Z:\\TWP\\KIBALI\\120108_Kibali_Fibre\\Instrumentation\\Documents\\NoArea\\NoRev\\120108-DGW120108-DWG-07GB-108-0810-01.dxf")
        print("DXF version: {}".format(dxf.dxfversion))
        header_var_count = len(dxf.header)
        layer_count = len(dxf.layers)
        entity_count = len(dxf.entities)
        blocks_count = len(dxf.blocks)
        ent = (dxf.entities)
        obj_ = (dxf.objects)
        blocks = (dxf.blocks)
        all_blocks = [entity for entity in dxf.entities if entity.dxftype == 'INSERT']
        print blocks_count
        # print layer_count

        # print all_blocks
        # print entity_count
        # print dxf.layers
        # for layer in dxf.layers:
        #	print layer.name, layer.color
        # for allblock in all_blocks:
        #	#print allblock
        #	print allblock.name
        #	print allblock.row_count
        #	print allblock.col_count
        #	attributes = (allblock.attribs)

        #	for a in attributes:
        #		print a.tag
        #		print a
        # print attdef(a)
        #	print "------"

        # for b in blocks:
        #	print b.name

        print obj_

    # Program seq starts here:
    ConnectODBC()
    CursorX()


class SplicePanel():
    pass


class SpliceTray():
    pass


class FibreCable():
    pass


if __name__ == '__main__':
    QButton()
