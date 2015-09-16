#! /usr/bin/env python2
# -*- coding: UTF-8 -*-
#
#  lfmsfiow
#  Last.fm scrobbler for iTunes on Windows
#
#  https://github.com/cvzi/Python/tree/master/lfmsfiow/
#

debug = False
import win32com.client
if win32com.client.gencache.is_readonly == True:

    win32com.client.gencache.is_readonly = False

    win32com.client.gencache.Rebuild()
from win32com.client.gencache import EnsureDispatch
import win32gui
import win32process
import pythoncom
import sys
import traceback
from time import sleep,time
import os
import threading
import multiprocessing
#import pylast
import pylast
import ConfigParser
# GUI stuff
from PyQt4 import QtGui,QtCore
import signal
import webbrowser
import urllib

# GUI?
gui = "pythonw" in sys.executable

def enum_callback(hwnd, data):
  # Get window that matches the processid pid
  pid = data[0]
  if pid == win32process.GetWindowThreadProcessId(hwnd)[1]:
    data[1] = hwnd

def echo(str,nobreak=False):
  if gui:
    return
  try:
    if nobreak:
      print str,
    else:
      print str
  except:
    try:
      str = ''.join([c if ord(c) < 128 else '?' for c in str])
      if nobreak:
        print str,
      else:
        print str
    except:
      traceback.print_exc(file=sys.stdout)

def clear():
  global debug
  global gui
  echo("\n")
  if not debug and not gui:
    os.system('cls')

doquit = False
g = None
scrobbleStatus = ["","","","",True] # [artist,title,album,statustext,changed]
class GUI(threading.Thread):
  class MainWindow:
    def __init__(self,w):
      self.window = w

      self.trayIcon = GUI.SystemTrayIcon(QtGui.QIcon("S.ico"), self)
      self.trayIcon.show()

    def app_exit(self):
      global doquit
      doquit = True
      self.window.close() # just close the main window

    def open_lastfm(self):
      global scrobbleStatus # [artist,title,album,statustext]
      webbrowser.open('http://www.last.fm/music/%s/%s/%s' % (urllib.quote_plus(scrobbleStatus[0]),urllib.quote_plus(scrobbleStatus[2]),urllib.quote_plus(scrobbleStatus[1])))

  class SystemTrayIcon(QtGui.QSystemTrayIcon):

    def __init__(self, icon, mw=None):
      global scrobbleStatus
      QtGui.QSystemTrayIcon.__init__(self, icon, mw.window)
      self.trayMenu = QtGui.QMenu(mw.window)

      self.statusEntry = QtGui.QAction(mw.window)
      self.statusEntry.setObjectName("statusEntry")
      self.statusEntry.setText(scrobbleStatus[3])
      self.trayMenu.addAction(self.statusEntry)
      QtCore.QObject.connect(self.statusEntry,QtCore.SIGNAL("triggered()"), mw.open_lastfm)

      self.actionExit = QtGui.QAction(mw.window)
      self.actionExit.setObjectName("actionExit")
      self.actionExit.setText("Quit")
      self.trayMenu.addAction(self.actionExit)
      QtCore.QObject.connect(self.actionExit,QtCore.SIGNAL("triggered()"), mw.app_exit)

      self.setContextMenu(self.trayMenu)

  def runinterpreter(self):
    global doquit
    global scrobbleStatus # [artist,title,album,statustext]

    self.mw.trayIcon.statusEntry.setText(QtCore.QString(scrobbleStatus[3]))
    self.mw.trayIcon.setToolTip(QtCore.QString(scrobbleStatus[3]))
    if scrobbleStatus[4]:
      qStr  = QtCore.QString("%s - %s" % (scrobbleStatus[0],scrobbleStatus[1]))
      self.mw.trayIcon.showMessage("Scrobbling", qStr)
      scrobbleStatus[4] = False

    if self.doquit or doquit:
      QtGui.QApplication.quit()


  def __init__(self):
    threading.Thread.__init__(self)
    self.doquit = False
    self.mw = None

  def run(self):
    self.app = QtGui.QApplication(sys.argv)

    timer = QtCore.QTimer()
    timer.start(500)
    timer.timeout.connect(self.runinterpreter)

    w = QtGui.QWidget()
    self.mw = self.MainWindow(w)

    sys.exit(self.app.exec_())

trackRepeated = False
trackRepeatedPlayCount = 0
trackRepeatedTrackId = []
class iTunesEventHandler:

  def OnAboutToPromptUserToQuitEvent(self):
    global doquit
    global g
    doquit = True
    if g:
      g.doquit = True

  def OnPlayerPlayEvent(self,track):
    global trackRepeated
    track = win32com.client.Dispatch(track)
    if trackRepeatedTrackId == track.GetITObjectIDs():
      if trackRepeatedPlayCount < track.PlayedCount:
        trackRepeated = True

  def OnPlayerStopEvent(self,track):
    global trackRepeated
    global trackRepeatedPlayCount
    global trackRepeatedTrackId
    track = win32com.client.Dispatch(track)
    trackRepeatedPlayCount = track.PlayedCount
    trackRepeatedTrackId = track.GetITObjectIDs()
    trackRepeated = False

class Scrobblethread(threading.Thread):
  def __init__(self):
    threading.Thread.__init__(self)
    self.doquit = False

  def run(self):
    global scrobblethreshold
    global scrobbleStatus # [artist,title,album,statustext]
    global trackRepeated

    scrobblethreshold = float(scrobblethreshold)
    pythoncom.CoInitialize()
    iTunes = EnsureDispatch("iTunes.Application")
    iTunesEvents = win32com.client.WithEvents(iTunes, iTunesEventHandler)
    echo("iTunes found!")
    i = 0
    startedat = 0
    lasttrack = None
    playing = 0 # Count seconds the same track was playing
    scrobbled = False
    while True:
      if self.doquit:
        del iTunes # delete all references to the com object, otherwise iTunes won't close because of an open API connection
        del iTunesEvents
        iTunes = None
        iTunesEvents = None
        return

      if iTunes.PlayerState == 1: # Playing
        t = iTunes.CurrentTrack
        track = (t.Artist,t.Name,t.Album,t.TrackNumber,t.Finish)
        if lasttrack != track or trackRepeated:
          # Started a new track
          trackRepeated = False
          lasttrack = track
          playing = 0
          scrobbled = False
          startedat = int(time())
          clear()
          #echo( "\n%s - %s (%s #%d) %ds" % (t.Artist,t.Name,t.Album,t.TrackNumber,t.Finish))
          echo( "%s - %s" % (t.Artist,t.Name))
          scrobbleStatus = [t.Artist,t.Name,t.Album,scrobbleStatus[3],True]

        scrobbleat = t.Finish * scrobblethreshold/100.0 # point to start scrobbling (in seconds)

        done = int(50.0/scrobbleat*playing)

        if not scrobbled and done >= 50:
          # scrobble
          network.scrobble(t.Artist,t.Name, int(time()), t.Album, track_number = t.TrackNumber, duration = t.Duration)
          scrobbled = True

        if scrobbled:
          done = 50
        undone = 50 - done;

        symbol = "-/|\\"[i]
        echo("\r%s%s%s%s" % ("#"*done , " " if scrobbled else symbol,  "o"*undone, " Scrobbled!" if scrobbled else ""),nobreak=True)

        if scrobbled:
          scrobbleStatus[3] = "%s - %s (Scrobbled!)" % (t.Artist,t.Name)
        else:
          scrobbleStatus[3] = "%s - %s (%d%%)" % (t.Artist,t.Name,int(done*2))

        i = (i+1) % 4
        playing += 1

      sleep(1)

# GUI?
if gui:
  g = GUI()

# Config
config = ConfigParser.RawConfigParser()
config.read('lfmsfiow.cfg')
API_KEY = config.get('Last.fm', 'apikey')
API_SECRET = config.get('Last.fm', 'apisecret')

scrobblethreshold = config.getfloat('User', 'scrobblethreshold')
username = config.get('User', 'username')
try:
  str = config.get('User', 'password')
  password_hash = pylast.md5(str)
except:
  password_hash = config.get('User', 'passwordhash')


# Get this window
window = [os.getpid(),None]
win32gui.EnumWindows(enum_callback,window)

# Last.fm
network = pylast.LastFMNetwork(api_key = API_KEY, api_secret =
    API_SECRET, username = username, password_hash = password_hash)

echo( "Connected to last.fm as %s" % username)

# iTunes
echo( "Waiting for iTunes ...")
thread = Scrobblethread()
thread.daemon = True
thread.start()

# Minimize console window
if not gui and window[1]:
  try:
    win32gui.CloseWindow(window[1])
  except:
    traceback.print_exc(file=sys.stdout)

# Infinite loop
try:
  if gui:
    g.daemon = True
    g.start()
  while True:
    sleep(1)
    if doquit:
      raise SystemExit()
except (KeyboardInterrupt, SystemExit) as e:

  echo( "")
  echo( "Ctrl-C quitting...")
  echo( "")
  thread.doquit = True
  sleep(1)
  del thread # delete all references to the com object, otherwise iTunes won't close because of an open API connection
  thread = None
  sleep(1) # Wait so everything is really closed
  if pythoncom._GetInterfaceCount() != 0:
    echo( "Not all references to the com object were cleared. %d references left." % pythoncom._GetInterfaceCount())
    sleep(5)
  pythoncom.CoUninitialize()