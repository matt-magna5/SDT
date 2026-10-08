MAGNA5 SDT - OFFLINE / AIR-GAPPED INSTALL
=========================================

Use this bundle on any box that cannot reach github.com, python.org or pypi.org.
Everything SDT needs is already inside it - nothing is downloaded.


INSTALL (the whole procedure)
-----------------------------
  1. Copy the zip to the target machine (USB, RDP drive redirection, file share).
  2. Right-click the zip -> Properties -> tick "Unblock" -> OK.      <-- do not skip
  3. Extract it (right-click -> Extract All). Extract it fully; do not run
     anything from inside the zip viewer.
  4. Double-click  RUN-SDT.cmd
  5. Open a NEW terminal and type:  sdt

Step 2 matters. Windows tags every file inside a downloaded zip with
Mark-of-the-Web; unblocking the zip first means the extracted files come out
clean. RUN-SDT.cmd also strips the tag itself, so skipping step 2 usually still
works - it is just one less thing to go wrong.

No admin rights are needed. SDT installs per-user into
%LOCALAPPDATA%\Magna5\SDT and adds its bin folder to your user PATH.


WHAT IS IN HERE
---------------
  deps\python-3.12.6-embed-amd64.zip   portable Python (never installed system-wide)
  deps\plink.exe                       SSH transport for ESXi / Linux hosts
  deps\wheels\*.whl                    pyVmomi, requests + dependencies
  deps\SHA256SUMS.txt                  checksums, verified during install
  VERSION                              the version this bundle was cut from

The installer verifies each dependency against SHA256SUMS.txt before using it.


OFFLINE BEHAVIOUR
-----------------
RUN-SDT.cmd sets SDT_OFFLINE=1 and passes -LocalSource, so the installer never
touches the network: no tag lookup, no release download, no pip. It also sets
SDT_NO_AUTOUPDATE=1 for your user so each `sdt` launch does not stall waiting
for a GitHub check that cannot succeed.

To re-enable updates later on a box that does have internet:
    setx SDT_NO_AUTOUPDATE ""
then open a new terminal and run:  sdt update


RUNNING A DISCOVERY
-------------------
    sdt            browser GUI (the normal way)
    sdt invoke     scan this box only, no UI
    sdt cli        console wizard
    sdt version    show version + install paths
    sdt uninstall  remove everything

Reports land in the session folder the GUI shows you, and on your Desktop under
SDT-Discovery-Sessions.


UPDATING AN OFFLINE BOX
-----------------------
There is no in-place offline update. Get a newer bundle, extract it somewhere
fresh, and run RUN-SDT.cmd again - it installs over the previous copy and keeps
your collected session data.


TROUBLESHOOTING
---------------
"running scripts is disabled on this system"
    You ran a .ps1 directly. Use RUN-SDT.cmd - it sets -ExecutionPolicy Bypass
    for that one process only and changes nothing on the machine.

"This IS the installed copy"
    You ran install.ps1 from %LOCALAPPDATA%\Magna5\SDT\app. Re-extract the
    bundle to a different folder (Desktop or Downloads) and run it from there.

"vSphere perf collection will be skipped"
    The wheels in deps\ did not verify or did not extract. Confirm deps\wheels
    survived the copy, then re-run RUN-SDT.cmd.

"sdt is not recognized"
    PATH changes only apply to NEW terminals. Open a fresh one. If it still
    fails, run:  %LOCALAPPDATA%\Magna5\SDT\bin\sdt.cmd
