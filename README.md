# OSCPoint Helper

A small Windows tray app that runs alongside [OSCPoint](https://oscpoint.com) on the PowerPoint machine and adds things OSCPoint doesn't do. It talks to PowerPoint directly (COM + UI Automation) and is controlled over OSC from the custom Companion module (`nisted/companion-module-zinc-oscpoint`, branch `nige-sync-status`).

**Status: 0.1 – untested on real PowerPoint.** Everything here needs checking on a Windows show machine.

## Features
| Feature | How |
|---|---|
| Start a **custom slide show** by name | PowerPoint object model (`SlideShowSettings`, restored afterwards) |
| **Go to build n** (no flash-through) | `SlideShowView.GotoClick` |
| **Scroll Presenter View notes** (buttons or rotary) | UI Automation scroll, falls back to mouse-wheel messages |
| **Notes font size** up/down | Presses Presenter View's own A+/A− buttons via UI Automation |
| **VT remaining-time overlay** | Always-on-top box, any size/position, turns red in the last N seconds |

## Install
Download `OSCPointHelper.exe` from the latest Actions run (or a Release), put it anywhere and run it. No install, no .NET needed. Add a shortcut to `shell:startup` to start it with Windows.

Settings and log: `%AppData%\OSCPointHelper\` (tray menu → Open settings / Open log).

## Ports
- Listens on UDP **35553** (`ListenPort`).
- Sends feedback to the IP that last sent it a command, on port **35550** (`ReplyPort`) – the module's normal OSCPoint feedback port, so no extra setup in Companion.

## OSC commands (all prefixed `/oscpoint/helper/`)
| Address | Args |
|---|---|
| `ping` | → replies `pong` (version) and `status` |
| `customshow/list` | → replies `customshows` (JSON array) |
| `customshow/start` | `s name` [, `i endRunning` 1/0, default 1] |
| `goto/build` | `i click` (0 = before first build) |
| `goto/slidebuild` | `i slide`, `i click` |
| `notes/scroll` | `i lines` (+down / −up) [, `i page` 1/0] |
| `notes/scroll/up`, `notes/scroll/down` | [`i lines`] |
| `notes/page/up`, `notes/page/down`, `notes/top`, `notes/bottom` | – |
| `notes/font/larger`, `notes/font/smaller` | [`i times`] |
| `notes/font` | `i steps` (+larger / −smaller) |
| `overlay/show`, `overlay/hide`, `overlay/toggle`, `overlay/larger`, `overlay/smaller` | – |
| `overlay/setup` | `i 1/0` – placeholder shown for positioning |

Feedback sent: `/oscpoint/helper/status` (`no-powerpoint` / `powerpoint` / `slideshow`, every 2 s), `/oscpoint/helper/customshows` (on change), `/oscpoint/helper/pong`.

## If notes scroll / font size don't work
These depend on Presenter View's internals, which Microsoft doesn't document. Tray menu → **Dump Presenter View controls** with Presenter View open, and send the text file – it shows exactly what PowerPoint exposes so the matching can be fixed.
