"""
Tenant Monitoring & Management — Web Edition (single-file build)
==================================================================
A phone-friendly web app that shares data with the original Tkinter desktop
app, synced entirely through the cloud service (see CLOUD_MODE below) --
there is no LAN/local-network pairing path in this build.

Run:
    pip install flask openpyxl reportlab qrcode[pil]
    python app.py

Deploy this file with CLOUD_MODE=1 and a DATABASE_URL, then pair phones to
it from the desktop app's Settings -> Connect Phone (which shows the QR
code for this deployed service directly).

Everything -- backend logic, the REST API, and the mobile frontend -- lives
in this one file on purpose, so it's a single thing to copy/share.
"""
import os
import re
import io
import copy
import json
import time
import shutil
import hashlib
import secrets
import calendar
import threading
import webbrowser
from urllib.parse import quote
from datetime import datetime, date, timezone

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                 Table, TableStyle, HRFlowable)

from flask import Flask, request, jsonify, session, send_file, Response, g
import base64

# Created here (rather than further down) because some decorators/routes
# below (e.g. @app.before_request for the CLOUD_MODE gate) reference `app`
# before the file reaches the "Flask app + REST API" section.
app = Flask(__name__)


@app.errorhandler(Exception)
def _api_error_handler(err):
    """Without this, any unhandled exception in a route (a bug, a bad
    Postgres value, whatever) falls through to Flask's default error
    page -- plain HTML, not JSON. The frontend's api() helper can't
    parse that as JSON, so it silently swallows the real cause and
    just shows the generic 'Request failed.' toast, which is exactly
    what makes bugs like this hard to track down from a bug report
    alone. Logging the real traceback server-side and returning it as
    JSON means the *next* time this fires, the toast (and the server
    log) shows what actually broke instead of a dead end."""
    import traceback
    traceback.print_exc()
    if request.path.startswith("/api/"):
        return jsonify({"ok": False, "error": f"{type(err).__name__}: {err}"}), 500
    raise err

APP_ICON_256_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAQAAAAEACAYAAABccqhmAAAXpElEQVR4nO3dbXCc13Uf8P+599nFvgEg9JKOUhaw6iqJ"
    "7Q+pTMNxh3ZAVHKhOpYYSgHCRIktUgIha1qPPP3QD84ERPop006tVlNVICKSjq2EXlgiKEW1MLUKwYniVBSHTib2xOV4"
    "KtCUM6UjwsAC+/Y8955+2F0SlPi2i8ULsf/fDL5wtItnhTln78u55wJERERERERERERERERERERERERERERERERERERE"
    "RERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERE"
    "RERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERERHVQlY1+BCLa"
    "CLXgHx01G/wkRLSuVA0A/PzTT98KAMhm7YY+DxGtC6kFe8+LRx6/czp7rjt7dAcAYGYm2NAno2viXI1WpzLkF4j4D7x4"
    "9CuIx56EAhBZ1ELxS3OD+w/3zcwEs/390UY/Kr0f52nUuNFRg8lJAxHf88LRQ6aj/Uktlr0vFj3CsMNkUs91Tx4Znu3v"
    "j3a8NR7b6Mel9+MIgBozOmowNuYBoOf40UM2kxl2C7kQggCAQNXDGjVtCeuKhQNnH9w3gWzWYmjIA9CNfXiqYQLYXAQ3"
    "Q3BUAtl1P/9Ml2Ta/9TE4wN+OR9Vg/8SVYUxatriRkvlp95+8JEvVRcKFSKb/3O2AE4BNovKtplu+u2zlcGfzkzbVHLA"
    "LS+/P/gBQETgvfhi0ZmO9id7Xjh6CCIek5Nm03/OFsERwGZQHU6333PPrbnXXnu3+q8GgN/Ix3qfavBvn5i4xd7e9qpJ"
    "tPW6peVQRK4/v1eEtrM95paWJub2PHIAwGXTCNoYTAAbra8vwOxs1Plrv7ZDE4mX1blv5qamvggAGBy0mJx0G/yEFTMz"
    "Afr7o+3HxnttJv28EXuXK+SdiLnxvX5FZNKpwJfL07qU+62zDz8xX0sqa/jkdA1MABupGvwdn/1sr8Zi00akC8bAR9G0"
    "E9mXP37872v/zYY+54rgD1KpabFBly/kHeoJ/ipVjYLOjsAt50/q8tIAk8DG4jxso7w3+IEujSKn5bIzQTAQqJ7q2L17"
    "ALOzUXW+vDHJuhr83ccmdgbJ1DQgDQc/AIhIEC0shiaR6JV0Zrr7a+N3YGjIsWpwYzABbIQrBb9zHiIWIlbLZQeRO8Ta"
    "V9t3736yOk9WDA6ua5DsGB+Pob8/6p48Mmwyqdeh2qXlsm80+GtEJOaXcs4EQa/Zljm1/dh4L4aGHKsG1x+nAOvt6sF/"
    "eTJW9RCBxOMGUZRFufz4wiuvzK/blECzFjLkuiePDNt08pAvFhXe6/uec1W/wztpS1ion4/y+YFze0dO1kYcTfsddE0c"
    "AaynGw1+ANV/M1oqRQiCIY3HX+v87Gc/itnZqDoSWJvkrVqp65ch94HjR0dMOnHIl4oO3qOpwQ8AYqyWig6QriCZmu4+"
    "NrET/f0RRwLrhwlgvdQT/CuJBFoqRWLM3RqPv965e/djF3cGmr2XriqYnDQYGnI9L371WUmmntVi0cF5A5G1SThirJbL"
    "HqpdJpN6vXvyyDD6+6Md4ywdXg+cAqyHRoN/JVUHY6wEAXwUjefm5/8NZmejpk0JRkcNDh5UiOil0t5FB8BijWL/Mqoe"
    "xohJJMQtFw6cHdw3UZuGrP0vb11MAGutGcF/iULVSzxu1bk3tFzel3vllTOrTgK1gpzR0aDnn3/wGZNJDfvFXAhgfb+F"
    "VRXWeGlrsyiVHn97zyPjyGYtBgc9S4fXBhPAWmpu8K8USRAE3vsLovqFxampbDWIFfWeJbi8tPdlk0js9EtLEUQam4er"
    "KgQekMZ2ClQVRrxJpa3PF8bnHvz840wCa4drAGtl7YIfAAINQyeqt4i132jfvftgQ1uF76nrN8nETpdbZfAbI5JIWkAb"
    "G7qLCLxav7TkbHt6pOf40UMXi4R4fqDpOAJYC2sb/JeoKgCVtjbjo2gai4sP51577d0bmhLUgv9r43dIe/KESSR6XS4X"
    "iTENBb+qdybeZjUsFxT4K5tK9ft8QaGKVSwghqajPeaX8hNz3/vRExgbi3h+oLmYUZttvYIfqASWiNFyOTLWDkhHx3cz"
    "99+/83rVgzvGx2MYGnLbj433mm2ZUyYIev1SzjUc/F4jm0pZQOfV496zD+3/l76Q/3eSaBOxttIboDExv7AYmWRiuOfu"
    "D77e/fwzXRgb86wabB4mgGZaz+C/XKDlshORu0wQvH7N6sGZmeDUyEi4/djh3iCVmgZwhy8WGi7tVdXIdrQH6vxpXSrd"
    "c3Zw/1/2zcwEcw/u/8+aLz6KWJCTWMyo+kanBIHLLUUmmdgp6cx09/PPdLF0uHmYAJqlOcEfNfxtKWI1ijy8tyYe/0rH"
    "nj3P3frAA+2YnHTo66t8s8+MVg/1HB4IEsE0IF1aKjYc/FBENp0OfLE47X+8cM/cbz92Gtmsna3u488N7j8c5fP3Qsy8"
    "TaWsqja0UyFGAreYi0yirVcyHdP/+Pn/fheTQHNwDaAZmhP8KrGYwHtoFLnqaxv5+ygAJ/F4oFF0Wpx7dOGll05/OJuN"
    "/2BoqFwr7dVSCQ2PTqor/baz07pc7tL5/vec6qs1A+35kz+6WzqSz0kQ3O2WliNpeJHRO5NMWXXuQpRfvu/c3pGTO8bH"
    "Y6dGRsKG3o+YAFZt9cFf2dqyVuDc70P105JIfEpLpUrxjzS8nRZJLBZ41XlRfWLx+PFj3d987oBNpcd9qejgvDQY/E6C"
    "wCIeg5bKX5rb8/mnoGpw8CCuuDhXW2x85pku2Z552cRXt81YW2wEdD7KhwPn9u7n+YFVYAJYjeYEv5cgsHBuZGFq6hD6"
    "+oKOrq4vQ+T3xJhAoygC0Fjtv6qDtRZhhPSvfux/tP3inZ9R57Va19/A+3lnkkmrzs2jVH787cFHs1A11f35q+/R10YG"
    "1UIj254edgsLDorGSoxVncTjFt7PR8Xot87t3T9dmd6MMQnUiWsAjWpO8KsEgUUYHliYmjqEwcE4Zmfd4tTUmDr3SfX+"
    "f0s8XuuyW98imkhlXaBY0tQnP6qJX/7QZzSMAO+lkaBTr5FJpqxXfybK5wfeHnw0u2N8PAaR63f5rQS/wR/8QTS353MH"
    "/GLuKZNIWhijDa15VI5Me6h2xTpSr1bOD4yx9XgDOAJoRHOC34m1AaLowMJLL01gx44YTp0KV74/+vqC9m3bfk+M+fKK"
    "0cD1h84igPfQMEJm18fR9pG7oMVSw+sKqhoGHR0xl8+fdD8t3XduePhCQ8PuFYeNul88MmwTiUNaKkMj52Eamo6w9fgq"
    "MQHUqzkHe0JJJmO+WBzLnThxEAcOxHDo0OULWSv6AbY/8MCviLX/RYLgV6rffFc/misCdQ5iLdI7P4q2D30QWigBpoE/"
    "tapCxNvODuvzhcNt54tP/vCxx3KrbuF1scXY4YEgEfypxGKNdxmqth6XtrjRcnhpTYKtx28IE0A9mrTVJ/F44MvlidyJ"
    "Eweq7+lw5W8sQV+frY0GOrq6vizGHAQAde79owEj0NDBxGNo/9e/iuCO21cT/F6MNYgH0LA8Nrdn30EAzevku7LJaCr9"
    "vLH2LpfPN1aJuCJRuYXcxNxDjxxANmvx/e8rqwavjQngRjUx+LVcnlg8ceLADR/gWTEa6Ni9ewDG/FcJgl+4bDRgBFoK"
    "IYk4Oj7Th+Af3QYtlgDT0MjamSBmEdjIl8tPXBxaN/tATi0JVNuM23S6N/rZQiSmwW1CILQdbD1eDyaAG7EWwV8J6nrm"
    "qhdHA1333tsZdXT8BwP8WwBQ772WQxPctg3pXZ9AcNs2aKncUPDDeTWZlGgYXfCF4gNn9w6/saZ77SsPJLVnnjVtiSG/"
    "tOSgDfYhUEQmkw58uZzV3OLjZx9+Yr66U8EkcAVMANezOYL/kveMBiQI/puWww/aWzu14zO7RBJxaClsbNjv1Zt0ShT+"
    "L8J88dF3fmP/mXVpyrHynsGprz5rEskRn1928NrQNqF6jYJtnYFbXmbr8etgAriWzRb8l1xaGwDkli8cON7+6Z0PwMKj"
    "7GyDc35Ios2H75wv5//81D3zX//6X2J8PIb1qrIbHTU4+BGp9SKUtsSzq6lWVNXQZtIxXyqd1HeWBs4+wSRwJUwAV7N5"
    "g7/yeNUy2+7jRx6y6cx/1ELxA1ouA6bO6K+sIUDiMRT/9ozLf/d7Vr37X7kdOz4N4MrVfWupGqQ9k4f3S7LtKTjf7sNy"
    "fTcQValqZNOpQJ07rYuFR+d++7HTLBi6HAuBrqRJPfzWKvh3vDUem63267fJ1De1ULyz4eA3BhILsPydk1iefRMSjwFB"
    "sHELZ9X7AS4dJJILNpmyaOA0oYgEfnnZiTF3Syb52vZj470sGLocE8B7Ne+b365B8AuyWXvqYyNh94vVfv35gtMw9HUH"
    "v1dIPAYNIyy+Movi356BtMUrSaHSaGTj9PdHfTMzwbmh4TfDXO4TPnInTXu7Vd/AaUIx1uXzDtCuIJWZ3n5svPfUx0ZC"
    "th6vYAJYqVlFPvF44KPo200NflWBjgqGhlzPC0cOmnj8kC8WPbw3dc+RvULaYnALOeS+9R2Ec+9AEtXg3yRm+/sjZLP2"
    "nYe/cEaXFgd8ofiGbc8EAOpekxAxVktlX0kCqel/MvncUC3JrMGj31SYAGqaE/yRxOMxH4YnAezF6KjB5GT9jTrfq/I+"
    "BjLme144eshu6xzVUsk1UNev8F4lEUf09z/F4tS3EZ1/F5JMAH7zBP9F1TP/Zx9+Yn7u9I92+UJxwm7rjAHq6h6lGDFa"
    "Lns47QraM9/onjwyPNvf3/LTgZbPgACaGfyBj6KTJgwHFl55ZR5TUwbA6ubSK7fIqv36o/mFUERidS3hVgJGTCaNwt/8"
    "ncv/+SlbW/yD38Rb5LWDRGNj0RxwoOeFIz+RtrZRLUce3tV3VZmIURd5zXu16eSh7heP4NTHWvv8AEcAaxX8lZ58q4us"
    "bNZibMx3P/9MV89LX/+GSaaG3UIuEqmzX7+ql1ggEguwfPJv/iH/3b+2EDhY0c007L+qsTFfu7Js7qF9B325fMAk2oxY"
    "a+DrPE0oYuC98cWiN/H4oZ4XjhzE0JCDjgpUW25XrLUTwFoG/2pX0d/Tstsmk0OVlt31jdpUvZNYzCAIck7x+fN/+NxH"
    "BPI/JZGwlZNDDbbvXm8iWtshOPvgvolwMX8fROZNKmnq3iEQEXgvWio5u61ztOeFo4cgYx6Tk6bVWo+31Ie9zE0Q/Nsn"
    "Jm6RdGbaJNp6o4WFsN4aefU+ssmUhciFaHn53h//+uf+GPnz5xemXvxXWi5/ESILEo9bABFuluFv9fLQc3v3T0fFwoB3"
    "7oxJpurfIRARQGw0vxDajszwxfsHKjcktUxctMwHvcxmDv6ZmaDWstvekfwrE4v1uqUlJyJ1DftVEQbbOgMfuZNhLveJ"
    "c0PDb2JmJoCqYHTULJ448bSqftw7Ny2xWFBdTLw5RgMXk8DISff/Cp/wYXgy6OwI0EDTURHE3EIuMonkcM+Jr73aaq3H"
    "Wy8BNLHIp9nBv2N8PFY7IhukMtNG7F2+kK+vCq6y2OdtRybmCsWsLi0OvPPwF84gm7XVBh6VI7J9fUFuaur/5I4fv0/D"
    "sDIaiMXi1SnB5h8NVLcJzw0PX0j8Q/keXyweNh3tQXWHoL73EgRueTmyqeRAq7Ueb60E0Kzgj8Wsd+6tZn/znxoZCbdn"
    "Jz5e6dev9TfJqNywqybRZnRp+am5B37nN696EKZ2eciK0YB6/y1pa7O4WXaHqjsEP3zssdzbD/zOo34pP27SGQsjDlrf"
    "AqyIBNHCYmgSbb2Szkxvn5i4pRWSQOusejYn+L1YawDMI4p2LLz88v9deTqv4Uertc9eUf+uYVjns3knNrDSFocrFivt"
    "sTRrgRs4w7/iKrHOPXv+vXr/qcUTJ+6vnhPY/KOBFa3GPnD86IgkEs9WWo1FdR8kUvXOJlPWqzvjlpYfPrd3ZEt3HW6N"
    "EUATg98D8xqGA80K/pV1/SaTeg5hVHfwq2pkUmkLI/PhYuG+sw/um8DMTAAZcjcUwCuuEls4fvwPJYp+F4DcFMEPXNoh"
    "yGbt23seGY+Wln8TRuYlHq/7RiIRY30h74zYu2qlw6hedLJWj7+Rtv4IoBqkzQp+CcOBxT/7s5PNCP5ao4pKg8zkoYb6"
    "9as6k0lbDaPTWiw8Ojf42OlVfWM143NtpBWtxoJU+lWx9haXL9TfZUi9k9r9A8XifeeGht/ciseJt/YIoFJCu7rgBxSV"
    "4L9wMfj7+oJVB8noaDX4D3/FtiUqdf2uzrp+VW9SSevLpTf8jxfuWXXwA6h+rpv3i6F2kGjvyMnaQSLb2RHUfS2ZmGrr"
    "cXQFqfS3eyb/aH+1YGhLxcyW+jCXqfbb67j//n+2yos6FUARpdL9F4P/eldvX0+twu+Fw/8p2Nb1pC8WGqjrV2dSSXGl"
    "0otzp36062LDi+bMVW+Oof9VvP8gUWHaptP1bxOKGA1Dr2HYbtrbn+v+5pGHIGN+cAstDG7dBPCDHwgAheo/NdZ2rbhv"
    "rx5aqR93+Ywx38PoqKl28F2dwe9XAkx1pxZLDhCtt/WVKhSxmIjTkxgbi3a89VZsqw1PV2VoyEHVnH34ifm53b97ny8W"
    "JmzXtto24Y0nuMrfP4QxKkAvAJy//fabd4T0Hls3AVSJteVq2+jG/2gishQEyWY3yRCRZQAWaLAGXRWApgBIJpe7qb+1"
    "14RIpapPR83cnkcO+IXFg9LWZsVYqfNGIkHlFEF+rR51o2z5BFA94LHqjG3CsOlH5rQZ//9v5GquVjY25iFjimzWvv3g"
    "I2O+XD6AeBBJEJh6z0Gobr142XIfiOgKFENDbsf4eOzsg/sm/FJ+F0TmTTJZaTWmWumHcM0fD4TOoHJR65bBBEAt49RI"
    "pRXY2b3Db0SF/IDXykEixGJeUgkg2XbVH0klYW7dVgDgfm7Xri0z4ro5Sj6JmqV2hmBo6OTPf/XpfxG7petbOP9ub/Tu"
    "BV+dFrzvJaoqUToFP/eTX7rtrg996vQXv/gmbpaDU9fBBECtp1o1+JOhoXd/8S+m7vnZ73/lO/qzxV+GgYdecafIQlUR"
    "j33OGLP3/Guv3QGghMra0k09GmACoNY0NOQwOGh/+Mlfz932Cx/6qU2nRF3kr7NerPCaE2tv6qBfiWsA1LomJxWqItbG"
    "4H2tJfrVfwABdEvFzJb6MER1E1HoZmyJvD6YAIhaGBMAUQtjAiBqYUwARC2MCYCohbEOgEjhqr0CrnVUWFEpEthSvQGZ"
    "AIgEXSYIAvUuuGYhkAg0im5X57ZMPwAmAGpltW/7p33k7lR1Xq9cCgwRURUVAy10dHQUFtfxIdcSEwC1MgWAn/7dD766"
    "mtffzJgAiAYHLc6fv/Fh/Wp7Qm4iTABEN3Mb9FXiNiBRC2MCIGphTABELYwJgKiFMQEQtTAmAKIWxgRA1MKYAIhaGBMA"
    "UQtjAiBqYUwARC2MCYCohTEBELUwJgCiFsYEQNTCmACIWhgTAFELYwIgamFMAEQtjAmAqIUxARC1MCYAohbGBEDUwpgA"
    "iFoYEwBRC2MCIGphTABELYwJgKiFbf3LQUUUQATAo/6EpwBEKq9vOgGcqkaqcIDWddW0QiNVhQj8WjwbtYYtnwBUJG7i"
    "8QAilVCulwi0XL5dw7CBF19Xl+3sCMSaAKa+3KTOBbazA9HSUmYNnotaxNZNAB/+sAKAOHdGy+UxDUNVqS8DiIiqqhig"
    "0N7ZWVhs2sMdVGAMInjaLyzeqaWiV61vdKLqvRhrYOR1ANi1a5efbdrzERHRlrcWw9rNRtDXZ1f9LrOzzV8HyGZt3+23"
    "r+pvMPv66x5jY1wHICIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi"
    "IiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi"
    "IiIiIiIiIiIiIiIiIqIm+/8GE8AfZ5jqXQAAAABJRU5ErkJggg=="
)


# =========================================================================
#  SECTION 1 -- Core data + business logic
#  (ported from the desktop app; framework-independent)
# =========================================================================
APP_NAME = "Tenant Monitoring & Management"

DATA_DIR   = os.path.join(os.path.expanduser("~"), ".rental_manager")
DATA_FILE  = os.path.join(DATA_DIR, "data.json")
PDF_FILE   = os.path.join(DATA_DIR, "tenant_data.pdf")
EXCEL_FILE = os.path.join(DATA_DIR, "tenant_records.xlsx")
BACKUP_DIR = os.path.join(DATA_DIR, "backups")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(BACKUP_DIR, exist_ok=True)

TENANT_FIELD_DEFAULTS = {
    "status":              "Pending",
    "payment_history":     list,
    "deposit_history":     list,
    "arrears_history":     list,
    "locked_periods":      list,
    "deposit_cycle_start": 0,
    "rent_increase_due":   0.0,
}

MONTH_PICKER_HORIZON = 24  # how many future open months to offer in the month picker

# ── cloud mode ───────────────────────────────────────────────────────────
# This exact same file can run two ways:
#   1) PC-local (default): can be run for local testing, reads/writes the
#      local ~/.rental_manager/data.json file directly.
#   2) Cloud (CLOUD_MODE=1 + DATABASE_URL set): deployed as its own
#      always-on Render web service, storing each paired install's data as
#      a row in Postgres instead of a local file, keyed by a session id.
#      Phones/browsers pair straight to this service and read/write it
#      directly, so tenant data stays reachable with the PC fully off.
# Nothing about the business logic below (routes, status/deposit-cycle
# calculations, etc.) changes between the two modes -- only where the
# `data` dict backing it comes from is swapped out.
CLOUD_MODE = os.environ.get("CLOUD_MODE") == "1"
_cloud_pool = None


def _parse_iso_dt(s):
    """Parses an ISO-8601 timestamp string into a timezone-AWARE datetime,
    treating a naive string (no offset) as UTC. Returns None if `s` is
    falsy or unparseable.

    This exists because timestamps compared for "which edit is newer"
    come from two different clocks: the PC's own `datetime.now()` (naive,
    in whatever timezone the PC's OS is set to) and Postgres's `now()`
    (UTC). Comparing those two ISO strings directly with `>=` -- which
    /api/_sync used to do -- is only valid when both strings use the same
    UTC offset. For a PC in a timezone ahead of UTC, its naive local-time
    string sorts as lexically LATER than a genuinely more recent UTC
    string, so every push looked artificially "current" and, worse, the
    PC's own bookkeeping of "the last cloud state I've seen" got stamped
    with that inflated local time -- silently masking every real edit
    made elsewhere for the rest of that UTC offset window. Parsing both
    into actual datetimes and comparing those (Python compares
    timezone-aware datetimes correctly, honoring the offset) fixes it."""
    if not s:
        return None
    try:
        dt = datetime.fromisoformat(str(s).replace("Z", "+00:00"))
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def _iso_ge(a, b):
    """True if timestamp string `a` is >= timestamp string `b`, comparing
    as actual datetimes (see _parse_iso_dt) rather than raw strings. A
    missing/unparseable `b` loses (anything real beats nothing); a
    missing/unparseable `a` never wins."""
    da, db = _parse_iso_dt(a), _parse_iso_dt(b)
    if da is None:
        return False
    if db is None:
        return True
    return da >= db

if CLOUD_MODE:
    import psycopg2
    import psycopg2.extras

    def _cloud_conn():
        global _cloud_pool
        db_url = os.environ.get("DATABASE_URL", "")
        if not db_url:
            raise RuntimeError("CLOUD_MODE is on but DATABASE_URL isn't set.")
        # A short-lived connection per request is plenty for a single
        # household's app and keeps this simple/robust across Render's
        # worker restarts -- no pool bookkeeping to get wrong.
        return psycopg2.connect(db_url)

    def _cloud_ensure_schema():
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cloud_sessions (
                    session_id   TEXT PRIMARY KEY,
                    secret_key   TEXT NOT NULL,
                    data         JSONB NOT NULL,
                    updated_at   TIMESTAMPTZ NOT NULL DEFAULT now(),
                    updated_by   TEXT
                )
            """)
            # Per-device roster for the cloud-direct pairing path -- the
            # counterpart to the LAN model's devices.json, just keyed by
            # session_id since one Postgres row backs every household's
            # phones instead of one file per install. last_seen/first_seen
            # are epoch floats (not TIMESTAMPTZ) so they drop straight into
            # _list_devices()-style "online = now - last_seen <= timeout"
            # math and the desktop's _format_last_synced() with no
            # conversion either side.
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cloud_devices (
                    session_id   TEXT NOT NULL,
                    device_id    TEXT NOT NULL,
                    label        TEXT,
                    first_seen   DOUBLE PRECISION NOT NULL,
                    last_seen    DOUBLE PRECISION NOT NULL,
                    kicked       BOOLEAN NOT NULL DEFAULT false,
                    fingerprint  TEXT,
                    PRIMARY KEY (session_id, device_id)
                )
            """)
            # fingerprint didn't exist on tables created before this change --
            # add it if missing so upgrades don't need a manual migration.
            cur.execute("""
                ALTER TABLE cloud_devices ADD COLUMN IF NOT EXISTS fingerprint TEXT
            """)
            # custom_label_locked: once a person names their own device (see
            # _cloud_set_device_label), the auto-detected user-agent label
            # must stop overwriting it on every subsequent poll.
            cur.execute("""
                ALTER TABLE cloud_devices ADD COLUMN IF NOT EXISTS custom_label_locked BOOLEAN NOT NULL DEFAULT false
            """)
            # The one-time pairing token that gates "/" -- see
            # _cloud_pairing_ok() below. A NULL token means nothing is
            # currently scannable; the desktop pushes a fresh one via
            # /api/pairing-token every time it (re)shows the QR code, and
            # it's cleared the instant it's consumed, same single-use
            # guarantee the LAN model already had.
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cloud_pairing (
                    session_id  TEXT PRIMARY KEY,
                    token       TEXT
                )
            """)
            # Holds the Flask session-signing secret itself. This used to
            # only ever live in a local devices.json file next to the
            # process -- fine on the desktop app, but on a Render (or
            # similar) web service the filesystem is ephemeral: every
            # restart/redeploy/free-tier spin-down wiped it, so a brand
            # new random secret got generated each time, silently
            # invalidating every phone's pairing cookie (see
            # _cloud_pairing_ok) and dropping them back to the "Waiting
            # to connect" screen even though nothing about their pairing
            # actually changed. Storing it here instead means it survives
            # restarts exactly like cloud_sessions/cloud_pairing already
            # do. Single fixed row (id=1) -- one secret for the whole
            # service, same as before.
            cur.execute("""
                CREATE TABLE IF NOT EXISTS cloud_app_secret (
                    id          INTEGER PRIMARY KEY,
                    secret_key  TEXT NOT NULL
                )
            """)
            conn.commit()

    _cloud_ensure_schema()

    def _cloud_get_or_create_secret():
        """Race-safe get-or-create: if two workers boot at once, only one
        INSERT wins and both end up reading the same row back."""
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                INSERT INTO cloud_app_secret (id, secret_key)
                VALUES (1, %s)
                ON CONFLICT (id) DO NOTHING
            """, (secrets.token_hex(32),))
            conn.commit()
            cur.execute("SELECT secret_key FROM cloud_app_secret WHERE id = 1")
            return cur.fetchone()[0]

    def _cloud_get_row(session_id):
        with _cloud_conn() as conn, conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT * FROM cloud_sessions WHERE session_id = %s", (session_id,))
            return cur.fetchone()

    def _cloud_load(session_id):
        row = _cloud_get_row(session_id)
        if not row:
            return {"units": {}, "tenants": [], "settings": {}}
        return row["data"]

    def _cloud_get_updated_at(session_id):
        """Cheap timestamp-only read (no `data` column, unlike
        _cloud_get_row) so lock_status can hand it to every device on
        every 6s ping -- see the JS pingServer() -- without paying to
        pull the full snapshot back down each time. Lets the web app
        notice a change (PC save, another phone, etc.) and apply it
        right away instead of waiting for its own slower full-refresh
        timer."""
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("SELECT updated_at FROM cloud_sessions WHERE session_id = %s", (session_id,))
            row = cur.fetchone()
            return row[0].isoformat() if row and row[0] else ""

    def _cloud_save(session_id, data, updated_by=None):
        """Always stamps updated_at with the DATABASE's own now() -- never
        a client-supplied value -- so it's authoritative regardless of
        which device's clock/timezone triggered the write. Returns that
        timestamp (ISO, UTC) so callers (notably /api/_sync below) can
        hand it back to whoever pushed, instead of them relying on their
        own local clock for bookkeeping."""
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                INSERT INTO cloud_sessions (session_id, secret_key, data, updated_at, updated_by)
                VALUES (%s, %s, %s, now(), %s)
                ON CONFLICT (session_id) DO UPDATE
                    SET data = EXCLUDED.data,
                        updated_at = now(),
                        updated_by = EXCLUDED.updated_by
                RETURNING updated_at
            """, (session_id, g.secret_key, json.dumps(data), updated_by))
            row = cur.fetchone()
            conn.commit()
            return row[0].isoformat() if row else None

    def _cloud_touch_device(session_id, device_id, user_agent=None, screen_hint="", fingerprint=""):
        """Cloud-mode counterpart to the LAN model's _touch_device(): same
        semantics (immediate admit, no separate approval step, label
        re-derived on every poll so it can upgrade from a generic
        'iPhone'/'Android' to a specific model, MAX_DEVICES cap per
        session) but backed by Postgres so the roster is shared by every
        phone hitting this session_id directly -- no PC involved -- and
        survives Render worker restarts.

        Returns (admitted, canonical_device_id). admitted is False only if
        the cap is full or it's kicked. canonical_device_id is normally
        just `device_id` echoed back -- except when `device_id` is brand
        new here but its `fingerprint` (stable hardware/browser signals,
        not tied to any one storage context -- see DEVICE_FINGERPRINT in
        the JS below) matches an ALREADY-known device on this roster. That
        happens whenever the same physical phone shows up through a
        storage context that doesn't share localStorage with wherever it
        was recognized before -- most commonly, iOS treats a web app
        "Added to Home Screen" as a separate storage silo from Safari
        itself, and some in-app/QR-scanner browsers do the same, even
        though it's the identical device. Rather than registering a
        second roster entry for what's really one phone, this reuses the
        existing device's row and hands its id back so the caller can
        tell the client to adopt it -- from then on that context IS that
        device, not a new one."""
        if not device_id:
            return True, device_id
        now = time.time()
        with _cloud_conn() as conn, conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(
                "SELECT kicked, custom_label_locked FROM cloud_devices WHERE session_id=%s AND device_id=%s",
                (session_id, device_id))
            rec = cur.fetchone()
            if rec is not None:
                if rec["kicked"]:
                    return False, device_id  # not admitted -- caller also checks this directly
                if rec["custom_label_locked"]:
                    # Person named this device themselves -- never let the
                    # auto-detected user-agent label clobber it again.
                    cur.execute(
                        "UPDATE cloud_devices SET last_seen=%s, fingerprint=COALESCE(NULLIF(%s,''), fingerprint) "
                        "WHERE session_id=%s AND device_id=%s",
                        (now, fingerprint, session_id, device_id))
                    conn.commit()
                    return True, device_id
                label = _label_for_user_agent(user_agent, screen_hint)
                if label and label != "Device":
                    cur.execute(
                        "UPDATE cloud_devices SET last_seen=%s, label=%s, fingerprint=COALESCE(NULLIF(%s,''), fingerprint) "
                        "WHERE session_id=%s AND device_id=%s",
                        (now, label, fingerprint, session_id, device_id))
                else:
                    cur.execute(
                        "UPDATE cloud_devices SET last_seen=%s, fingerprint=COALESCE(NULLIF(%s,''), fingerprint) "
                        "WHERE session_id=%s AND device_id=%s",
                        (now, fingerprint, session_id, device_id))
                conn.commit()
                return True, device_id

            # Unknown device_id -- before registering it as a new phone,
            # see if an already-known, non-kicked device on this roster
            # has the same fingerprint. If so, this is the same physical
            # phone showing up from a different storage context.
            if fingerprint:
                cur.execute(
                    "SELECT device_id, custom_label_locked FROM cloud_devices "
                    "WHERE session_id=%s AND fingerprint=%s AND kicked=false "
                    "ORDER BY last_seen DESC LIMIT 1",
                    (session_id, fingerprint))
                match = cur.fetchone()
                if match:
                    canonical_id = match["device_id"]
                    if match["custom_label_locked"]:
                        cur.execute(
                            "UPDATE cloud_devices SET last_seen=%s "
                            "WHERE session_id=%s AND device_id=%s",
                            (now, session_id, canonical_id))
                        conn.commit()
                        return True, canonical_id
                    label = _label_for_user_agent(user_agent, screen_hint)
                    if label and label != "Device":
                        cur.execute(
                            "UPDATE cloud_devices SET last_seen=%s, label=%s "
                            "WHERE session_id=%s AND device_id=%s",
                            (now, label, session_id, canonical_id))
                    else:
                        cur.execute(
                            "UPDATE cloud_devices SET last_seen=%s "
                            "WHERE session_id=%s AND device_id=%s",
                            (now, session_id, canonical_id))
                    conn.commit()
                    return True, canonical_id

            cur.execute(
                "SELECT COUNT(*) AS n FROM cloud_devices WHERE session_id=%s AND kicked=false",
                (session_id,))
            if cur.fetchone()["n"] >= MAX_DEVICES:
                return False, device_id
            cur.execute("""
                INSERT INTO cloud_devices (session_id, device_id, label, first_seen, last_seen, kicked, fingerprint)
                VALUES (%s, %s, %s, %s, %s, false, %s)
                ON CONFLICT (session_id, device_id) DO NOTHING
            """, (session_id, device_id, _label_for_user_agent(user_agent, screen_hint), now, now, fingerprint))
            conn.commit()
            return True, device_id

    def _cloud_device_kicked(session_id, device_id):
        if not device_id:
            return False
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute(
                "SELECT kicked FROM cloud_devices WHERE session_id=%s AND device_id=%s",
                (session_id, device_id))
            row = cur.fetchone()
            return bool(row and row[0])

    def _cloud_list_devices(session_id):
        """Snapshot for the desktop app's admin list -- same shape as the
        LAN model's _list_devices() so _render_connected_devices_list()
        on the desktop side needs no changes: device_id, short_id, label,
        last_seen (epoch), online."""
        now = time.time()
        with _cloud_conn() as conn, conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("""
                SELECT device_id, label, last_seen FROM cloud_devices
                WHERE session_id=%s AND kicked=false
                ORDER BY last_seen DESC
            """, (session_id,))
            rows = cur.fetchall()
        return [
            {
                "device_id": r["device_id"],
                "short_id": r["device_id"][:8],
                "label": r["label"] or "Unknown device",
                "last_seen": r["last_seen"],
                "online": (now - r["last_seen"]) <= ONLINE_TIMEOUT,
            }
            for r in rows
        ]

    def _cloud_kick_device(session_id, device_id):
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                INSERT INTO cloud_devices (session_id, device_id, label, first_seen, last_seen, kicked)
                VALUES (%s, %s, 'Unknown device', %s, %s, true)
                ON CONFLICT (session_id, device_id) DO UPDATE SET kicked=true
            """, (session_id, device_id, time.time(), time.time()))
            conn.commit()

    def _cloud_set_device_label(session_id, device_id, label):
        """Lets a phone give itself a custom name (e.g. 'Mary's iPhone')
        instead of the generic auto-detected model name -- and locks it so
        _cloud_touch_device never overwrites it again on a later poll."""
        now = time.time()
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                INSERT INTO cloud_devices (session_id, device_id, label, first_seen, last_seen, kicked, custom_label_locked)
                VALUES (%s, %s, %s, %s, %s, false, true)
                ON CONFLICT (session_id, device_id)
                DO UPDATE SET label=%s, custom_label_locked=true, last_seen=%s
            """, (session_id, device_id, label, now, now, label, now))
            conn.commit()

    def _cloud_active_device_count(session_id):
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute(
                "SELECT COUNT(*) FROM cloud_devices WHERE session_id=%s AND kicked=false",
                (session_id,))
            return cur.fetchone()[0]

    def _cloud_set_pairing_token(session_id, token):
        """Desktop calls this every time it (re)shows the QR code --
        immediately replaces whatever token was active before, so an old
        screenshot or a link copied from an earlier QR code stops working
        the moment a new one is shown, not just when it's actually used."""
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                INSERT INTO cloud_pairing (session_id, token) VALUES (%s, %s)
                ON CONFLICT (session_id) DO UPDATE SET token = EXCLUDED.token
            """, (session_id, token))
            conn.commit()

    def _cloud_check_and_consume_pairing_token(session_id, token):
        """Single-use, race-safe: the UPDATE only matches (and clears) the
        token if it's still exactly what's stored, so two near-simultaneous
        requests for the same forwarded link can't both succeed -- only
        whichever one reaches Postgres first."""
        if not token:
            return False
        with _cloud_conn() as conn, conn.cursor() as cur:
            cur.execute("""
                UPDATE cloud_pairing SET token = NULL
                WHERE session_id = %s AND token = %s
            """, (session_id, token))
            matched = cur.rowcount > 0
            conn.commit()
            return matched

    def _cloud_pairing_ok(session_id, token):
        """Cloud-mode counterpart of the LAN model's _pairing_ok(): '/'
        may only bootstrap the app shell for (a) a browser that already
        paired successfully before -- a persistent cookie, so reopening
        an already-installed PWA doesn't demand a fresh scan every time
        -- or (b) a request presenting the current, unconsumed token. A
        plain forwarded link (sid+key with no valid token, from a
        browser that never scanned) gets neither."""
        cookie_key = "cloud_paired_" + session_id
        if session.get(cookie_key):
            return True
        if _cloud_check_and_consume_pairing_token(session_id, token):
            session[cookie_key] = True
            session.permanent = True
            return True
        return False


@app.before_request
def _cloud_gate():
    """Only active in CLOUD_MODE. Resolves which household's data this
    request is for (X-Session-Id / X-Secret-Key headers) and blocks the
    handful of routes that only make sense running on the actual PC
    (device pairing/kicking, PIN lock, shutdown, LAN QR code, etc.)."""
    if not CLOUD_MODE:
        return None

    if request.method == "OPTIONS":
        # A browser's CORS preflight for a genuine cross-origin request
        # (see cloudFetch() in the companion's JS -- reached when a page
        # is running locally with no direct-cloud pairing and falls back
        # to this service's own origin) never carries the real
        # X-Session-Id/X-Secret-Key headers, only lists them in
        # Access-Control-Request-Headers. Gating this like a normal
        # request would 401 the preflight itself, which makes the
        # browser block the real request entirely -- silently, with no
        # visible error, just a fall-through to stale local cache. Let
        # it through here; _cors_headers below decorates the response.
        return None

    path = request.path
    pc_only_prefixes = (
        "/api/device-count",
        "/api/announce-disconnect",
        "/api/unlock", "/api/lock", "/api/settings/pin",
        "/api/settings/reset", "/api/shutdown", "/connect", "/qr.png",
        "/api/cloud-config",
    )
    # /api/lock-status, /api/devices(/…/kick), and /api/pairing-token used
    # to be blanket-404'd here along with the rest of this list -- they
    # made sense only for the PC-hosted LAN companion. Now that cloud-
    # direct pairing has its own device roster (cloud_devices) and its
    # own single-use scan token (cloud_pairing), they're real, session-
    # authenticated endpoints instead: /api/lock-status is what a cloud-
    # direct phone's pingServer() polls to know the cloud service itself
    # is reachable, /api/devices is how the desktop manages that roster,
    # and /api/pairing-token is how the desktop pushes a fresh scan token
    # every time it (re)shows the QR code -- see _cloud_pairing_ok().
    if path.startswith(pc_only_prefixes):
        return Response("", status=404, mimetype="text/plain")
    if path in ("/manifest.json", "/sw.js") or path.startswith("/icon-"):
        # These carry no tenant data, so there's nothing to gate -- let
        # their own routes below answer regardless of mode.
        return None
    if path == "/":
        # Handled entirely by index() below (it checks ?sid=&key= itself
        # in CLOUD_MODE), so a phone that scanned the direct-cloud QR code
        # (see get_direct_cloud_pairing_url()) can load the app shell
        # straight from this service with the PC not involved at all.
        return None

    if not path.startswith("/api/"):
        return None

    session_id = request.headers.get("X-Session-Id", "")
    secret_key = request.headers.get("X-Secret-Key", "")
    if not session_id or not secret_key:
        return jsonify({"ok": False, "error": "session_required"}), 401

    if path == "/api/_sync" and request.method in ("PUT", "GET"):
        # A session with no row yet isn't a bad secret -- it just hasn't
        # been created by a first push yet (fresh cloud database, or a
        # phone/browser reaching this service before the PC's first
        # push has landed). PUT already treated this as "fine, this
        # push establishes it"; GET used to fall through to the generic
        # check below and get a hard bad_secret 403 instead of reaching
        # cloud_sync()'s own, already-correct {"exists": false} response
        # -- which is exactly what looked like "cloud unreachable" on
        # both the phone (blank page) and the PC (device list stuck on
        # "retrying...") even though the service was up the whole time.
        row = _cloud_get_row(session_id)
        if row and row["secret_key"] != secret_key:
            return jsonify({"ok": False, "error": "bad_secret"}), 403
        g.session_id, g.secret_key = session_id, secret_key
        return None

    row = _cloud_get_row(session_id)
    if not row or row["secret_key"] != secret_key:
        return jsonify({"ok": False, "error": "bad_secret"}), 403
    g.session_id, g.secret_key = session_id, secret_key

    # Presence tracking for the cloud device roster -- every authenticated
    # request from a phone (not the PC's own push above, which never sends
    # X-Device-Id) refreshes its last-seen time here, not just its polls
    # to /api/lock-status, so "last synced" reflects real activity.
    device_id = request.headers.get("X-Device-Id", "")
    if device_id:
        if _cloud_device_kicked(session_id, device_id):
            return jsonify({"ok": False, "error": "kicked"}), 403
        admitted, canonical_id = _cloud_touch_device(
            session_id, device_id,
            request.headers.get("User-Agent"),
            request.headers.get("X-Device-Screen", ""),
            request.headers.get("X-Device-Fingerprint", ""))
        if not admitted:
            return jsonify({"ok": False, "error": "device_limit_reached",
                             "max_devices": MAX_DEVICES}), 403
        if canonical_id != device_id:
            g.canonical_device_id = canonical_id
    return None


@app.after_request
def _cors_headers(resp):
    """CLOUD_MODE only: cloudFetch() in the companion's JS is a genuine
    cross-origin request whenever a page is running locally with no
    direct-cloud pairing and falls back to reading/writing this cloud
    service's own origin for durability. Without these headers the
    browser silently blocks it (no visible error -- api() just falls
    through to a stale local cache), which is exactly the "shows old
    data instead of live cloud data while offline" symptom this fixes.
    A wildcard origin is safe here: these routes authenticate via the
    explicit X-Session-Id / X-Secret-Key headers above, never via
    cookies, so there's no session/cookie for a hostile origin to ride
    along on."""
    if CLOUD_MODE:
        resp.headers["Access-Control-Allow-Origin"] = "*"
        resp.headers["Access-Control-Allow-Headers"] = (
            "Content-Type, X-Session-Id, X-Secret-Key, X-Device-Id, X-Device-Fingerprint")
        resp.headers["Access-Control-Allow-Methods"] = "GET, POST, PUT, DELETE, OPTIONS"
        canonical_id = getattr(g, "canonical_device_id", None)
        if canonical_id:
            # Tells the client its current X-Device-Id wasn't recognized as
            # new, but matched an already-known device's fingerprint (see
            # _cloud_touch_device) -- the client adopts this id from here
            # on instead of staying registered as a separate device.
            resp.headers["X-Canonical-Device-Id"] = canonical_id
            resp.headers["Access-Control-Expose-Headers"] = "X-Canonical-Device-Id"
    return resp


# ── low level helpers ───────────────────────────────────────────────────
def parse_amount(value):
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = re.sub(r"[^\d.\-]", "", str(value))
    if s in ("", "-", ".", "-."):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def add_months(d, n):
    n = int(n)
    total_month_index = (d.month - 1) + n
    year = d.year + total_month_index // 12
    month = total_month_index % 12 + 1
    last_day = calendar.monthrange(year, month)[1]
    day = min(d.day, last_day)
    return d.replace(year=year, month=month, day=day)


def calculate_one_month_ahead(date_string):
    try:
        d = datetime.strptime(date_string.strip(), "%Y-%m-%d").date()
        return add_months(d, 1).strftime("%Y-%m-%d")
    except (ValueError, AttributeError):
        return ""


def hash_secret(secret):
    return hashlib.sha256(secret.encode()).hexdigest()


def _parse_date(s):
    if not s:
        return None
    try:
        return datetime.strptime(s.strip(), "%Y-%m-%d").date()
    except (ValueError, AttributeError):
        return None


def _compute_txn_period(rec):
    to_d = _parse_date(rec.get("to_date") or "")
    pay_d = _parse_date(rec.get("date") or "")
    if to_d is not None:
        from_d = add_months(to_d, -1)
        if pay_d is not None and not (from_d <= pay_d <= to_d):
            to_d = pay_d
            from_d = add_months(to_d, -1)
    elif pay_d is not None:
        to_d = pay_d
        from_d = add_months(to_d, -1)
    else:
        return "—", "—"
    return from_d.strftime("%Y-%m-%d"), to_d.strftime("%Y-%m-%d")


def fmt_period_date(d):
    """'24 jul,26' -- day, lowercase abbreviated month, comma, 2-digit year."""
    return f"{d.day} {d.strftime('%b').lower()},{d.strftime('%y')}"


def fmt_period(from_str, to_str):
    """'24 jul,26 to 24 aug,26' from two YYYY-MM-DD strings."""
    f, t = _parse_date(from_str), _parse_date(to_str)
    if not f or not t:
        return "—"
    return f"{fmt_period_date(f)} to {fmt_period_date(t)}"


def open_months_list(t, horizon=MONTH_PICKER_HORIZON):
    """The still-payable months for tenant t, starting at the first unpaid
    month (the tenant's current due date, or their move-in date if they've
    never paid), running forward `horizon` months. Each entry's `months`
    is how many consecutive months would be paid if the picker were
    confirmed with that entry as the last one ticked."""
    anchor = _parse_date(t.get("due_date", "")) or _parse_date(t.get("entry_date", "")) or date.today()
    out = []
    d = anchor
    for i in range(horizon):
        nxt = add_months(d, 1)
        out.append({
            "months": i + 1,
            "from": d.strftime("%Y-%m-%d"),
            "to": nxt.strftime("%Y-%m-%d"),
            "label": fmt_period(d.strftime("%Y-%m-%d"), nxt.strftime("%Y-%m-%d")),
        })
        d = nxt
    return out


def cleared_months_list(t):
    """The already-paid months, from the tenant's move-in month up to (not
    including) their current due date -- shown in the month picker as
    locked/checked-off history, since a cleared month can't be selected
    again."""
    entry = _parse_date(t.get("entry_date", ""))
    due = _parse_date(t.get("due_date", ""))
    out = []
    if not entry or not due or due <= entry:
        return out
    d = entry
    while d < due:
        nxt = add_months(d, 1)
        if nxt > due:
            nxt = due
        out.append({
            "from": d.strftime("%Y-%m-%d"),
            "to": nxt.strftime("%Y-%m-%d"),
            "label": fmt_period(d.strftime("%Y-%m-%d"), nxt.strftime("%Y-%m-%d")),
        })
        d = nxt
    return out


# ── data load / save ────────────────────────────────────────────────────
def _unique_backup_path(prefix, ext):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    candidate = os.path.join(BACKUP_DIR, f"{prefix}_{ts}{ext}")
    n = 1
    while os.path.exists(candidate):
        candidate = os.path.join(BACKUP_DIR, f"{prefix}_{ts}_{n}{ext}")
        n += 1
    return candidate


def backup_current_data_file():
    if not os.path.exists(DATA_FILE):
        return None
    backup_path = _unique_backup_path("data", ".json")
    shutil.copy2(DATA_FILE, backup_path)
    return backup_path


def _raw_load():
    if CLOUD_MODE:
        return _cloud_load(g.session_id)
    if not os.path.exists(DATA_FILE):
        return {"units": {}, "tenants": [], "settings": {}}
    try:
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        try:
            corrupt_copy = _unique_backup_path("data_CORRUPT", ".json")
            shutil.copy2(DATA_FILE, corrupt_copy)
        except Exception:
            pass
        return {"units": {}, "tenants": [], "settings": {}}


def _record_content_equal(a, b):
    """True if two tenant/unit records are the same ignoring the
    `_updated_at` bookkeeping field itself (added by
    _stamp_changed_records below), so re-saving an unchanged record
    never looks like an edit. Mirrors the identically-named helper in
    the desktop app (updates.py) so both sides stamp records the same
    way."""
    a2 = {k: v for k, v in (a or {}).items() if k != "_updated_at"}
    b2 = {k: v for k, v in (b or {}).items() if k != "_updated_at"}
    return a2 == b2


def _tenant_key(t):
    return (t.get("name"), t.get("unit"), t.get("entry_date"))


def _tenant_key_str(key):
    """String form of a _tenant_key() tuple, safe to use as a JSON object
    key (tombstone dicts are stored inside data.json, which requires
    string keys). Uses a separator that can't appear in a name/unit/date
    field entered through the UI."""
    return "\x1f".join("" if p is None else str(p) for p in key)


def stamp_tenant_deleted(data, t):
    """Record a tombstone for a tenant that's about to be permanently
    removed from data["tenants"], so a later cloud merge (_merge_cloud_data)
    knows this record was deliberately deleted here rather than treating
    the OTHER side's still-existing copy of it as something new that must
    be kept -- which is exactly what silently resurrected deleted tenants
    before this existed. Call this BEFORE removing the record from the
    list, and before save_raw()/save_state()."""
    tomb = data.setdefault("deleted_tenants", {})
    tomb[_tenant_key_str(_tenant_key(t))] = datetime.now(timezone.utc).isoformat()


def stamp_unit_deleted(data, name):
    """Record a tombstone for a unit that's about to be permanently
    removed from data["units"]. See stamp_tenant_deleted() above -- same
    reasoning, just keyed by unit name (already a plain string)."""
    tomb = data.setdefault("deleted_units", {})
    tomb[name] = datetime.now(timezone.utc).isoformat()


def _stamp_changed_records(new_data, old_data):
    """Stamps `_updated_at` (ISO now) onto every tenant/unit record in
    new_data that's brand new or whose content actually changed versus
    old_data; unchanged records keep whatever `_updated_at` they
    already had. Mutates new_data in place and returns it.

    Mirrors the desktop app's identically-named helper -- this is what
    lets the PC's cloud-merge logic (_merge_cloud_data in updates.py)
    tell, record by record, whether its own copy or the phone's is
    actually newer, instead of only being able to compare whole
    snapshots (which is what used to let one side's edit to a shared
    record get silently discarded whenever the other side's snapshot
    was picked as "the base")."""
    now_iso = datetime.now(timezone.utc).isoformat()
    old_tenants = {_tenant_key(t): t for t in (old_data.get("tenants") or [])}
    for t in (new_data.get("tenants") or []):
        old_t = old_tenants.get(_tenant_key(t))
        if old_t is None or not _record_content_equal(t, old_t):
            t["_updated_at"] = now_iso
        elif "_updated_at" not in t and old_t.get("_updated_at"):
            t["_updated_at"] = old_t["_updated_at"]

    old_units = old_data.get("units") or {}
    for uk, uv in (new_data.get("units") or {}).items():
        if not isinstance(uv, dict):
            continue
        old_u = old_units.get(uk)
        if old_u is None or not _record_content_equal(uv, old_u):
            uv["_updated_at"] = now_iso
        elif "_updated_at" not in uv and isinstance(old_u, dict) and old_u.get("_updated_at"):
            uv["_updated_at"] = old_u["_updated_at"]

    # Mirrors the desktop app's settings stamping -- see the identically
    # named function in updates.py for why this exists.
    new_settings = new_data.get("settings")
    if isinstance(new_settings, dict):
        old_settings = old_data.get("settings") if isinstance(old_data.get("settings"), dict) else {}
        if not _record_content_equal(new_settings, old_settings):
            new_settings["_updated_at"] = now_iso
        elif "_updated_at" not in new_settings and old_settings.get("_updated_at"):
            new_settings["_updated_at"] = old_settings["_updated_at"]

    return new_data


def _merge_cloud_data(a, b):
    """Record-level, order-independent merge of two full snapshots (`a`
    and `b` can be passed in either order -- the result is the same
    either way). Ported 1:1 from the desktop app's identically-named
    helper so the server can do the same safe merge the PC already does
    client-side, instead of the PUT /api/_sync handler picking an
    entire snapshot as "the winner" by comparing the pushing side's own
    wall-clock timestamp against the row's stored (server-clock)
    timestamp -- that comparison is what silently dropped every PC push
    once its clock drifted even a few seconds behind the database's,
    since accepted writes are always stamped with the DB's own now().

    For each tenant/unit key:
      - present in only one side -> keep it.
      - present in both -> keep whichever copy has the newer
        `_updated_at` stamp (see _stamp_changed_records).
      - a record missing `_updated_at` entirely is treated as maximally
        stale; if neither side has a stamp, `a`'s copy wins (arbitrary
        but stable).
    `settings` is treated as a single block and resolved by its own
    `_updated_at` stamp -- whichever side's settings were actually
    edited more recently wins outright, rather than `a` always winning
    regardless of recency.

    `deleted_tenants` / `deleted_units` are tombstones: {key: deleted_at
    iso timestamp}, written by stamp_tenant_deleted()/stamp_unit_deleted()
    at the moment a record is permanently removed on either side (see
    those functions). Without this, a record present on only one side
    was always just "kept" -- indistinguishable from a genuinely NEW
    record the other side hadn't seen yet -- which is what silently
    resurrected a tenant/unit deleted on one device the next time it
    merged against the other device's still-intact copy. The fix: merge
    the tombstones themselves (newest timestamp per key wins, union of
    both sides), then drop any record whose key has a tombstone dated at
    or after that record's own `_updated_at` (a record with no stamp at
    all is treated as older than any tombstone, same "maximally stale"
    rule used elsewhere in this function) -- so a genuine edit made
    AFTER the deletion (e.g. the record was deleted, then re-added fresh
    with the same name/unit) still wins, but a stale copy that predates
    the deletion does not come back from the dead."""
    def _tenants_by_key(d):
        return {(t.get("name"), t.get("unit"), t.get("entry_date")): t
                for t in (d.get("tenants") or [])}

    a_del_t, b_del_t = (a.get("deleted_tenants") or {}), (b.get("deleted_tenants") or {})
    merged_del_t = dict(a_del_t)
    for k, ts in b_del_t.items():
        if k not in merged_del_t or _iso_ge(ts, merged_del_t[k]):
            merged_del_t[k] = ts

    a_del_u, b_del_u = (a.get("deleted_units") or {}), (b.get("deleted_units") or {})
    merged_del_u = dict(a_del_u)
    for k, ts in b_del_u.items():
        if k not in merged_del_u or _iso_ge(ts, merged_del_u[k]):
            merged_del_u[k] = ts

    def _tombstoned(record, tomb_ts):
        return tomb_ts is not None and _iso_ge(tomb_ts, record.get("_updated_at", ""))

    a_tenants, b_tenants = _tenants_by_key(a), _tenants_by_key(b)
    merged_tenants = []
    for key, ta in a_tenants.items():
        tb = b_tenants.get(key)
        winner = (ta if tb is None or _iso_ge(ta.get("_updated_at", ""), tb.get("_updated_at", "")) else tb)
        if _tombstoned(winner, merged_del_t.get(_tenant_key_str(key))):
            continue
        merged_tenants.append(winner)
    for key, tb in b_tenants.items():
        if key not in a_tenants:
            if _tombstoned(tb, merged_del_t.get(_tenant_key_str(key))):
                continue
            merged_tenants.append(tb)

    a_units, b_units = (a.get("units") or {}), (b.get("units") or {})
    merged_units = {}
    for uk, ua in a_units.items():
        ub = b_units.get(uk)
        if not isinstance(ua, dict):
            winner = ub if ub is not None else ua
        elif not isinstance(ub, dict):
            winner = ua
        else:
            winner = (ua if _iso_ge(ua.get("_updated_at", ""), ub.get("_updated_at", "")) else ub)
        if isinstance(winner, dict) and _tombstoned(winner, merged_del_u.get(uk)):
            continue
        merged_units[uk] = winner
    for uk, ub in b_units.items():
        if uk not in a_units:
            if isinstance(ub, dict) and _tombstoned(ub, merged_del_u.get(uk)):
                continue
            merged_units[uk] = ub

    a_settings, b_settings = (a.get("settings") or {}), (b.get("settings") or {})
    merged_settings = (
        a_settings if not b_settings or
        _iso_ge(a_settings.get("_updated_at", ""), b_settings.get("_updated_at", ""))
        else b_settings)

    return {
        "tenants": merged_tenants,
        "units": merged_units,
        "settings": merged_settings,
        "deleted_tenants": merged_del_t,
        "deleted_units": merged_del_u,
    }


def save_raw(data, updated_by=None):
    try:
        old_data = _raw_load()
    except Exception:
        old_data = {"units": {}, "tenants": [], "settings": {}}
    _stamp_changed_records(data, old_data)
    if CLOUD_MODE:
        return _cloud_save(g.session_id, data, updated_by=updated_by or "cloud")
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    # Same fix as the desktop app's save_data(): export_excel/export_pdf
    # rebuild the whole workbook/document from scratch, which is real
    # work -- doing it synchronously here would hold the HTTP response
    # (and so the phone's UI) hostage until both finished, on every
    # single save. The JSON write above is what everything else actually
    # depends on, so the response can go back now; exports finish a
    # moment later in the background. Deepcopy for the same reason as
    # the desktop side: `data` may still be referenced/mutated by the
    # caller after this function returns.
    export_snapshot = copy.deepcopy(data)

    def _export_worker():
        try:
            export_excel(export_snapshot)
        except Exception:
            pass
        try:
            export_pdf(export_snapshot)
        except Exception:
            pass

    threading.Thread(target=_export_worker, daemon=True).start()
    return None


def load_state():
    """Load data.json and run the same startup migrations the desktop app
    runs, returning a plain dict with 'units', 'tenants', 'settings'."""
    data = _raw_load()
    data.setdefault("units", {})
    data.setdefault("tenants", [])
    data.setdefault("settings", {})
    changed = False

    # normalize rent fields
    for t in data["tenants"]:
        raw_rent = t.get("rent", 0)
        if not isinstance(raw_rent, (int, float)):
            t["rent"] = parse_amount(raw_rent)
            changed = True
    for _name, info in data["units"].items():
        if isinstance(info, dict):
            raw_rent = info.get("rent", 0)
            if not isinstance(raw_rent, (int, float)):
                info["rent"] = parse_amount(raw_rent)
                changed = True

    # backfill schema fields
    for t in data["tenants"]:
        for field, default in TENANT_FIELD_DEFAULTS.items():
            if field not in t:
                t[field] = default() if callable(default) else default
                changed = True

    # apply due rent increases
    today = date.today()
    for unit_name, info in data["units"].items():
        if not isinstance(info, dict):
            continue
        pending = info.get("pending_rent_increase")
        if not pending:
            continue
        eff = _parse_date(str(pending.get("effective_month", "")))
        if eff is None:
            info.pop("pending_rent_increase", None)
            changed = True
            continue
        if eff <= today:
            new_rent = parse_amount(pending.get("new_rent", 0))
            info["rent"] = new_rent
            info.pop("pending_rent_increase", None)
            changed = True
            for t in data["tenants"]:
                if t.get("unit") == unit_name:
                    old_tenant_rent = parse_amount(t.get("rent", 0))
                    _bill_rent_increase_shortfall(t, old_tenant_rent, new_rent, eff)
                    t["rent"] = new_rent

    # auto revert expired "Confirmed" tenants back to Pending
    for t in data["tenants"]:
        if t.get("status") != "Confirmed":
            continue
        due = _parse_date(t.get("due_date", ""))
        if due and due < today:
            t["status"] = "Pending"
            t["pay_date"] = ""
            changed = True

    # deposit-in-progress tenants whose current month is already covered
    # should show as Confirmed, mirroring desktop app startup pass
    for t in data["tenants"]:
        if t.get("status") == "Confirmed":
            continue
        _, _, cleared, _ = current_deposit_cycle(t)
        if cleared:
            t["status"] = "Confirmed"
            if not t.get("pay_date"):
                t["pay_date"] = today.strftime("%Y-%m-%d")
            changed = True

    if changed:
        save_raw(data)
    return data


def save_state(data):
    save_raw(data)


# ── tenant status / deposit-cycle logic (ported 1:1 from desktop app) ──
def current_deposit_cycle(t):
    rent_target = parse_amount(t.get("rent", 0))
    full_history = t.get("deposit_history", [])
    cycle_start = t.get("deposit_cycle_start", 0)
    if not isinstance(cycle_start, int) or cycle_start < 0 or cycle_start > len(full_history):
        cycle_start = 0
    cycle_window = full_history[cycle_start:]
    active_records = [r for r in cycle_window if not r.get("_cancelled")]
    cancelled_records = [r for r in cycle_window if r.get("_cancelled")]

    paid_so_far = sum(float(r.get("amount", 0)) for r in active_records)
    current_bal = max(0.0, rent_target - paid_so_far)

    if cancelled_records:
        cancelled_sum = sum(float(r.get("amount", 0)) for r in cancelled_records)
        remaining = min(current_bal + cancelled_sum, rent_target)
        paid_so_far = max(0.0, rent_target - remaining)
        cleared = remaining <= 0
        in_progress = bool(active_records) and not cleared
    else:
        remaining = current_bal
        cleared = paid_so_far >= rent_target and rent_target > 0
        in_progress = bool(active_records) and not cleared
    return paid_so_far, remaining, cleared, in_progress


def deposit_paid_so_far(t):
    """Sum of active (non-cancelled) deposits since the last time the
    deposit cycle cleared. All deposits in this window are toward the same
    open month-window by construction (there's a single running window
    now, not separate current/next/multiple buckets), so this no longer
    needs to filter by a period label."""
    full_history = t.get("deposit_history", [])
    cycle_start = t.get("deposit_cycle_start", 0)
    if not isinstance(cycle_start, int) or cycle_start < 0 or cycle_start > len(full_history):
        cycle_start = 0
    cycle_window = full_history[cycle_start:]
    return sum(float(r.get("amount", 0)) for r in cycle_window if not r.get("_cancelled"))


def is_current_period_paid(t):
    if t.get("status") != "Confirmed":
        return False
    due = _parse_date(t.get("due_date", ""))
    if due is None:
        return False
    return due >= date.today()


def snapshot_tenant_state(t):
    return {
        "due_date": t.get("due_date", ""),
        "status": t.get("status", "Pending"),
        "pay_date": t.get("pay_date", ""),
        "locked_periods": list(t.get("locked_periods", [])),
        "deposit_cycle_start": t.get("deposit_cycle_start", 0),
        "rent_increase_due": t.get("rent_increase_due", 0),
    }


def restore_tenant_state(t, snap):
    t["due_date"] = snap.get("due_date", "")
    t["status"] = snap.get("status", "Pending")
    if snap.get("pay_date"):
        t["pay_date"] = snap["pay_date"]
    else:
        t.pop("pay_date", None)
    if snap.get("locked_periods"):
        t["locked_periods"] = list(snap["locked_periods"])
    else:
        t.pop("locked_periods", None)
    t["deposit_cycle_start"] = snap.get("deposit_cycle_start", 0)
    t["rent_increase_due"] = snap.get("rent_increase_due", 0)


def has_prior_payment_history(t):
    return bool(t.get("payment_history")) or bool(t.get("deposit_history"))


def due_date_shift_base(t, has_prior_history=None):
    if has_prior_history is None:
        has_prior_history = has_prior_payment_history(t)
    base = t.get("due_date", "") if has_prior_history else t.get("entry_date", "")
    return base or date.today().strftime("%Y-%m-%d")


def pending_reference_date_str(t):
    return t.get("due_date", "") or t.get("entry_date", "")


def status_level(t, today):
    """Returns (level, label) where level is 'paid' | 'underpaid' | 'pending'."""
    due = _parse_date(t.get("due_date", ""))
    if t.get("status") == "Confirmed":
        if due is not None and due < today:
            return "pending", "Pending"
        return "paid", "Paid"
    _, _, dep_cleared, dep_in_progress = current_deposit_cycle(t)
    if dep_cleared:
        if due is not None and due < today:
            return "pending", "Pending"
        return "paid", "Paid"
    if dep_in_progress:
        return "underpaid", "Installments"
    return "pending", "Pending"


def rent_increase_due(t):
    try:
        return max(0.0, float(t.get("rent_increase_due", 0) or 0))
    except (TypeError, ValueError):
        return 0.0


def _bill_rent_increase_shortfall(t, old_rent, new_rent, effective):
    if new_rent <= old_rent:
        return
    due_d = _parse_date(t.get("due_date", ""))
    if due_d is None:
        return
    months_prepaid_past_effective = (
        (due_d.year - effective.year) * 12 + (due_d.month - effective.month)
    )
    if months_prepaid_past_effective <= 0:
        return
    shortfall = (new_rent - old_rent) * months_prepaid_past_effective
    if shortfall > 0:
        t["rent_increase_due"] = float(t.get("rent_increase_due", 0) or 0) + shortfall


def record_arrears_payment(t, amount, method):
    pay_date = date.today().strftime("%Y-%m-%d")
    t.setdefault("arrears_history", []).append({
        "date": pay_date, "amount": amount, "method": method,
        "type": "Arrears", "txn_id": secrets.token_hex(4),
    })
    t["rent_increase_due"] = max(0.0, rent_increase_due(t) - amount)


# ── excess cascade (non-interactive web version) ────────────────────────
def apply_excess_cascade(t, excess, rent_target, base_due_str, pay_date, txn_id=None):
    """Web-simplified version of the desktop app's interactive cascade: a
    month is auto-recorded as a Full Payment whenever the remaining excess
    covers it completely, otherwise as a partial Deposit toward that month
    (matching the same rule the desktop dialog uses to decide whether
    "Full" is even offered). Returns the number of months the due date
    was advanced by."""
    if excess is None or excess <= 0 or rent_target is None or rent_target <= 0:
        return 0
    cur_due = _parse_date(base_due_str)
    if cur_due is None:
        return 0

    remaining = excess
    last_cleared = False
    full_months_shift = 0
    while remaining > 0.01:
        next_due = add_months(cur_due, 1)
        chunk = min(remaining, rent_target)
        allow_full = chunk >= rent_target - 0.01

        if allow_full:
            t.setdefault("payment_history", []).append({
                "date": pay_date, "months": 1,
                "amount": chunk,
                "from_date": cur_due.strftime("%Y-%m-%d"),
                "to_date": next_due.strftime("%Y-%m-%d"),
                "txn_id": txn_id,
            })
            cur_due = next_due
            t["due_date"] = cur_due.strftime("%Y-%m-%d")
            t["status"] = "Confirmed"
            t["pay_date"] = pay_date
            full_months_shift += 1
            last_cleared = True
            t["deposit_cycle_start"] = len(t.get("deposit_history", []))
        else:
            t.setdefault("deposit_history", []).append({
                "date": pay_date, "months": 1,
                "amount": chunk, "txn_id": txn_id,
                "from_date": cur_due.strftime("%Y-%m-%d"),
                "to_date": next_due.strftime("%Y-%m-%d"),
                "target_month": next_due.strftime("%Y-%m-%d"),
            })
            t["deposit_cycle_start"] = len(t["deposit_history"]) - 1
            last_cleared = False
        remaining -= chunk

    t["status"] = "Confirmed" if last_cleared else "Pending"
    return full_months_shift


# ── payments / deposits ─────────────────────────────────────────────────
def record_payment(t, months):
    """Mutates tenant dict t in place. `months` is the number of
    consecutive open months ticked in the month picker (1 = just the next
    due month, n = that month plus the n-1 months after it). Returns dict
    describing the result."""
    months = max(1, int(months))

    rent = parse_amount(t.get("rent", 0))
    total = rent * months

    pre_txn_state = snapshot_tenant_state(t)
    txn_id = secrets.token_hex(4)

    old_due_str = t.get("due_date", "")
    shift_base_str = due_date_shift_base(t)
    new_due_str = old_due_str
    old_due = _parse_date(shift_base_str)
    if old_due:
        new_due_str = add_months(old_due, months).strftime("%Y-%m-%d")

    pay_date = date.today().strftime("%Y-%m-%d")
    t["pay_date"] = pay_date
    t["status"] = "Confirmed"
    if new_due_str:
        t["due_date"] = new_due_str

    record_to_date = new_due_str
    record_from_date = shift_base_str if record_to_date else ""

    t.setdefault("payment_history", []).append({
        "date": pay_date, "months": months, "amount": total,
        "from_date": record_from_date, "to_date": record_to_date,
        "txn_id": txn_id, "_pre_state": pre_txn_state,
    })

    pre_paid, _, _, _ = current_deposit_cycle(t)
    if pre_paid > 0:
        t["deposit_cycle_start"] = len(t.get("deposit_history", []))
        apply_excess_cascade(t, pre_paid, rent, t["due_date"], pay_date, txn_id=txn_id)

    return {"amount": total, "months": months,
            "due_date": t["due_date"], "old_due_date": old_due_str,
            "period_label": fmt_period(record_from_date, record_to_date)}


def record_deposit(t, months, instalment):
    """`months` is the number of consecutive open months ticked in the
    month picker; the instalment accumulates toward rent x months."""
    months = max(1, int(months))

    rent_target = parse_amount(t.get("rent", 0))
    dep_paid_so_far = deposit_paid_so_far(t)
    effective_target = rent_target * months
    balance_before = max(0.0, effective_target - dep_paid_so_far)
    excess = max(0.0, instalment - balance_before)
    applied_amount = instalment - excess
    new_balance = max(0.0, effective_target - dep_paid_so_far - applied_amount)
    pay_date = date.today().strftime("%Y-%m-%d")

    had_prior_history = has_prior_payment_history(t)
    pre_txn_state = snapshot_tenant_state(t)
    txn_id = secrets.token_hex(4)

    shift_base_str = due_date_shift_base(t, had_prior_history)
    target_to_str = ""
    shift_base_d = _parse_date(shift_base_str)
    if shift_base_d:
        target_to_str = add_months(shift_base_d, months).strftime("%Y-%m-%d")

    t.setdefault("deposit_history", []).append({
        "date": pay_date, "months": months, "amount": applied_amount,
        "from_date": shift_base_str, "to_date": target_to_str,
        "txn_id": txn_id, "_pre_state": pre_txn_state,
    })

    if new_balance <= 0:
        prev_start = t.get("deposit_cycle_start", 0)
        t["deposit_history"][-1]["_cycle_start_before_clear"] = prev_start
        t["deposit_cycle_start"] = len(t["deposit_history"])

        t["due_date"] = target_to_str or t.get("due_date", "")
        t["status"] = "Confirmed"
        t["pay_date"] = pay_date

    if excess > 0.01:
        apply_excess_cascade(t, excess, rent_target, t.get("due_date") or shift_base_str,
                              pay_date, txn_id=txn_id)

    return {"amount": applied_amount, "excess": excess, "new_balance": new_balance,
            "cleared": new_balance <= 0, "due_date": t.get("due_date", ""),
            "period_label": fmt_period(shift_base_str, target_to_str)}


def cancel_transaction(t, h_key, idx):
    """Cancel a payment/deposit record at t[h_key][idx], reversing every
    linked record that shares its txn_id (or falling back to legacy single
    record logic for records saved before txn_id existed)."""
    history = t.get(h_key, [])
    if idx < 0 or idx >= len(history):
        return None
    rec = history[idx]
    if rec.get("_cancelled"):
        return "already_cancelled"
    txn_id = rec.get("txn_id")

    linked = []
    if txn_id:
        for key in ("payment_history", "deposit_history"):
            for i, r in enumerate(t.get(key, [])):
                if r.get("txn_id") == txn_id and not r.get("_cancelled"):
                    linked.append((key, i, r))
    else:
        linked = [(h_key, idx, rec)]

    cancelled_on = date.today().strftime("%Y-%m-%d")
    pre_state = None
    for key, i, r in linked:
        entry = t.get(key, [])[i]
        entry["_cancelled"] = True
        entry["_type_origin"] = "deposit" if key == "deposit_history" else "payment"
        entry["cancelled_on"] = cancelled_on
        if "_pre_state" in entry and pre_state is None:
            pre_state = entry["_pre_state"]

    if pre_state is not None:
        restore_tenant_state(t, pre_state)
    else:
        # Legacy record saved before _pre_state existed -- best effort.
        if h_key == "deposit_history":
            cycle_start = t.get("deposit_cycle_start", 0)
            if isinstance(cycle_start, int) and cycle_start > idx:
                prev_start = rec.get("_cycle_start_before_clear", idx)
                t["deposit_cycle_start"] = prev_start
        t["status"] = "Pending"
        t.pop("pay_date", None)
        t.pop("locked_periods", None)

    total_amt = sum(float(r.get("amount", 0)) for _, _, r in linked)
    return {"n_records": len(linked), "total_amount": total_amt}


MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]


def build_monthly_report(data, year, month):
    """Aggregates payment/deposit history for one calendar month across all
    tenants -- the same "Monthly Report" the desktop app's Reports screen
    generates, which had no equivalent here yet. Returns per-tenant summary
    rows, a flat list of individual transactions for the detail table, and
    the month's grand totals (active vs cancelled)."""
    tenants = data.get("tenants", [])
    prefix = f"{year:04d}-{month:02d}"
    month_label = f"{MONTHS[month - 1]} {year}"

    tenant_rows = []
    detail_rows = []
    grand_pay = grand_dep = grand_cancelled = 0

    for t in tenants:
        t_name = t.get("name", "—")
        t_unit = t.get("unit", "—")
        pay_active = pay_cancelled = 0
        dep_active = dep_cancelled = 0

        for rec in t.get("payment_history", []):
            if not rec.get("date", "").startswith(prefix):
                continue
            amt = int(rec.get("amount", 0))
            is_c = rec.get("_cancelled", False)
            ft, tt = _compute_txn_period(rec)
            detail_rows.append({
                "name": t_name, "unit": t_unit, "date": rec.get("date", "—"),
                "txn_type": "Full Payment" + (" (Cancelled)" if is_c else ""),
                "amount": amt, "from_d": ft, "to_d": tt,
                "is_cancelled": is_c, "cancelled_on": rec.get("cancelled_on", "—"),
            })
            if is_c:
                pay_cancelled += amt
            else:
                pay_active += amt

        for rec in t.get("deposit_history", []):
            if not rec.get("date", "").startswith(prefix):
                continue
            amt = float(rec.get("amount", 0))
            is_c = rec.get("_cancelled", False)
            ft, tt = _compute_txn_period(rec)
            detail_rows.append({
                "name": t_name, "unit": t_unit, "date": rec.get("date", "—"),
                "txn_type": "Deposit" + (" (Cancelled)" if is_c else ""),
                "amount": int(amt), "from_d": ft, "to_d": tt,
                "is_cancelled": is_c, "cancelled_on": rec.get("cancelled_on", "—"),
            })
            if is_c:
                dep_cancelled += amt
            else:
                dep_active += amt

        grand_pay += pay_active
        grand_dep += dep_active
        grand_cancelled += pay_cancelled + dep_cancelled

        if pay_active + dep_active + pay_cancelled + dep_cancelled > 0:
            tenant_rows.append({
                "name": t_name, "unit": t_unit,
                "pay_active": pay_active, "dep_active": dep_active,
                "pay_cancelled": pay_cancelled, "dep_cancelled": dep_cancelled,
            })

    detail_rows.sort(key=lambda r: (r["name"], r["date"]))
    return {
        "prefix": prefix,
        "month_label": month_label,
        "tenant_rows": tenant_rows,
        "detail_rows": detail_rows,
        "grand_pay": grand_pay,
        "grand_dep": grand_dep,
        "grand_cancelled": grand_cancelled,
        "grand_combined": grand_pay + grand_dep,
    }


def export_monthly_excel(report):
    """Writes a single "Monthly Transactions" sheet for one month, styled to
    match the desktop app's dedicated monthly export (separate from the
    all-time Transaction History workbook produced by export_excel())."""
    C_WHITE, C_GREY_H, C_CANCEL, C_TOTAL = "FFFFFF", "DDDDDD", "FFF0F0", "F0F7FF"

    def side():
        s = Side(style="thin", color="000000")
        return Border(left=s, right=s, top=s, bottom=s)

    def bold(sz=11):
        return Font(name="Calibri", bold=True, size=sz, color="000000")

    def reg(sz=10):
        return Font(name="Calibri", size=sz, color="000000")

    def strike(sz=10):
        return Font(name="Calibri", size=sz, color="000000", strike=True)

    def ctr():
        return Alignment(horizontal="center", vertical="center", wrap_text=True)

    def lft():
        return Alignment(horizontal="left", vertical="center", wrap_text=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Monthly — {report['month_label']}"[:31]
    ws.sheet_view.showGridLines = True

    ws.merge_cells("A1:I1")
    title = ws["A1"]
    title.value = f"Monthly Transactions — {report['month_label']}"
    title.alignment = ctr()
    title.fill = PatternFill("solid", fgColor="1A73E8")
    title.font = Font(name="Calibri", bold=True, size=13, color="FFFFFF")
    ws.row_dimensions[1].height = 30

    hdrs = ["#", "TENANT", "UNIT", "DATE", "TYPE", "AMOUNT (UGX)", "FROM", "TO", "NOTE"]
    widths = [5, 24, 10, 14, 18, 20, 14, 14, 22]
    for ci, (h, cw) in enumerate(zip(hdrs, widths), 1):
        cell = ws.cell(row=2, column=ci, value=h)
        cell.font = bold(10)
        cell.fill = PatternFill("solid", fgColor=C_GREY_H)
        cell.alignment = ctr()
        cell.border = side()
        ws.column_dimensions[get_column_letter(ci)].width = cw
    ws.row_dimensions[2].height = 22

    row_num = 3
    for seq, dr in enumerate(report["detail_rows"], 1):
        is_c = dr["is_cancelled"]
        fgc = C_WHITE if not is_c else C_CANCEL
        note = f"Cancelled on {dr['cancelled_on']}" if is_c else ""
        row_vals = [
            (seq, ctr()), (dr["name"], lft()), (dr["unit"], ctr()),
            (dr["date"], ctr()), (dr["txn_type"], ctr()), (dr["amount"], ctr()),
            (dr["from_d"], ctr()), (dr["to_d"], ctr()), (note, lft()),
        ]
        for ci, (val, aln) in enumerate(row_vals, 1):
            cell = ws.cell(row=row_num, column=ci, value=val)
            cell.fill = PatternFill("solid", fgColor=fgc)
            cell.border = side()
            cell.alignment = aln
            cell.font = strike(10) if is_c else (bold(10) if ci in (2, 6) else reg(10))
        ws.row_dimensions[row_num].height = 20
        row_num += 1

    if row_num == 3:
        ws.cell(row=3, column=1, value="No transactions recorded this month.").font = reg(11)
        row_num += 1

    ws.merge_cells(f"A{row_num}:E{row_num}")
    tc = ws.cell(row=row_num, column=1, value="TOTALS")
    tc.font, tc.alignment, tc.border = bold(11), ctr(), side()
    tc.fill = PatternFill("solid", fgColor=C_TOTAL)
    for ci in range(2, 6):
        c2 = ws.cell(row=row_num, column=ci)
        c2.fill, c2.border = PatternFill("solid", fgColor=C_TOTAL), side()
    pay_cell = ws.cell(row=row_num, column=6,
                        value=f"Pay: {report['grand_pay']:,}  Dep: {int(report['grand_dep']):,}  "
                              f"Cancelled: {int(report['grand_cancelled']):,}")
    pay_cell.font, pay_cell.alignment, pay_cell.border = bold(10), ctr(), side()
    pay_cell.fill = PatternFill("solid", fgColor=C_TOTAL)
    for ci in range(7, 10):
        c3 = ws.cell(row=row_num, column=ci, value="")
        c3.fill, c3.border = PatternFill("solid", fgColor=C_TOTAL), side()
    ws.row_dimensions[row_num].height = 22

    monthly_path = os.path.join(DATA_DIR, f"monthly_{report['prefix'].replace('-', '_')}.xlsx")
    wb.save(monthly_path)
    return monthly_path


# ── exports (ported near-verbatim from desktop app) ─────────────────────
def export_excel(data):
    wb = openpyxl.Workbook()
    C_WHITE = "FFFFFF"
    C_BLACK = "000000"

    def thin_border():
        s = Side(style="thin", color="000000")
        return Border(left=s, right=s, top=s, bottom=s)

    def bold(sz=11):
        return Font(name="Calibri", bold=True, size=sz, color=C_BLACK)

    def reg(sz=10):
        return Font(name="Calibri", size=sz, color=C_BLACK)

    def cancelled_font(sz=10):
        return Font(name="Calibri", size=sz, color=C_BLACK, strike=True)

    def center():
        return Alignment(horizontal="center", vertical="center", wrap_text=True)

    def left():
        return Alignment(horizontal="left", vertical="center", wrap_text=True)

    tenants = data.get("tenants", [])
    ws = wb.active
    ws.title = "Transaction History"
    ws.sheet_view.showGridLines = True

    hdrs = ["#", "TENANT NAME", "UNIT", "DATE OF PAYMENT", "TYPE OF PAYMENT",
            "AMOUNT (UGX)", "FROM", "TO"]
    col_widths = [5, 24, 10, 18, 20, 22, 16, 16]
    for ci, (h, cw) in enumerate(zip(hdrs, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = bold(10)
        cell.fill = PatternFill("solid", fgColor="DDDDDD")
        cell.alignment = center()
        cell.border = thin_border()
        ws.column_dimensions[get_column_letter(ci)].width = cw
    ws.row_dimensions[1].height = 24

    row_num = 2
    seq = 0
    for t in tenants:
        all_txns = []
        for rec in t.get("payment_history", []):
            all_txns.append({"_type": "payment", **rec})
        for rec in t.get("deposit_history", []):
            all_txns.append({"_type": "deposit", **rec})
        all_txns.sort(key=lambda r: r.get("date", ""), reverse=True)

        for rec in all_txns:
            seq += 1
            is_cancelled = rec.get("_cancelled", False)
            if rec["_type"] == "deposit":
                type_str = "Deposit" + (" (Cancelled)" if is_cancelled else "")
            else:
                type_str = "Full Payment" + (" (Cancelled)" if is_cancelled else "")
            try:
                amt_str = f"{int(rec.get('amount', 0)):,}"
            except Exception:
                amt_str = "—"
            ft, tt = _compute_txn_period(rec)
            row_vals = [
                (seq, center()), (t.get("name", "—"), left()), (t.get("unit", "—"), center()),
                (rec.get("date", "—"), center()), (type_str, center()),
                (amt_str, center()), (ft, center()), (tt, center()),
            ]
            for ci, (val, aln) in enumerate(row_vals, 1):
                cell = ws.cell(row=row_num, column=ci, value=val)
                cell.fill = PatternFill("solid", fgColor=C_WHITE)
                cell.border = thin_border()
                cell.alignment = aln
                if is_cancelled:
                    cell.font = cancelled_font(10)
                else:
                    cell.font = bold(10) if ci in (2, 6) else reg(10)
            ws.row_dimensions[row_num].height = 20
            row_num += 1

    if row_num == 2:
        ws.cell(row=2, column=1, value="No transactions recorded yet.").font = reg(11)

    wb.save(EXCEL_FILE)
    return EXCEL_FILE


def export_pdf(data):
    doc = SimpleDocTemplate(
        PDF_FILE, pagesize=A4,
        leftMargin=18 * mm, rightMargin=18 * mm, topMargin=20 * mm, bottomMargin=20 * mm,
    )
    W = A4[0] - 36 * mm
    styles = getSampleStyleSheet()
    H1 = ParagraphStyle("H1", parent=styles["Heading1"], fontSize=18,
                         textColor=colors.HexColor("#0E4F4F"), spaceAfter=4)
    H2 = ParagraphStyle("H2", parent=styles["Heading2"], fontSize=13,
                         textColor=colors.HexColor("#1FAD9F"), spaceAfter=2, spaceBefore=10)
    MUTED = ParagraphStyle("MUTED", parent=styles["Normal"], fontSize=9,
                            textColor=colors.HexColor("#6E8482"))
    BOLD = ParagraphStyle("BOLD", parent=styles["Normal"], fontSize=10, fontName="Helvetica-Bold")

    story = []
    today_str = date.today().strftime("%d %B %Y")
    story.append(Spacer(1, 10 * mm))
    story.append(Paragraph(APP_NAME, H1))
    story.append(Paragraph(f"Data Export — {today_str}", MUTED))
    story.append(HRFlowable(width=W, thickness=1, color=colors.HexColor("#1FAD9F"), spaceAfter=8))

    tenants = data.get("tenants", [])
    units = data.get("units", {})
    total_collected = int(
        sum(int(rec.get("amount", 0)) for t in tenants for rec in t.get("payment_history", [])
            if not rec.get("_cancelled", False))
        + sum(float(rec.get("amount", 0)) for t in tenants for rec in t.get("deposit_history", [])
              if not rec.get("_cancelled", False))
    )
    today = date.today()

    def _days_status(t):
        ref_str = t.get("due_date", "") or t.get("entry_date", "")
        due = _parse_date(ref_str)
        if due is None:
            return "—"
        rem = (due - today).days
        if t.get("status") == "Confirmed" and rem >= 0:
            return f"Paid  |  {rem}d until next due"
        return f"{'Overdue by ' + str(abs(rem)) + 'd' if rem < 0 else str(rem) + 'd remaining'}"

    summary_data = [
        ["Total Tenants", "Units", "Total Income", "Export Date"],
        [str(len(tenants)), str(len(units)), f"UGX {total_collected:,}", today_str],
    ]
    ts = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0E4F4F")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#EAF7F4")),
        ("FONTSIZE", (0, 1), (-1, 1), 11),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#1FAD9F")),
        ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#B0D4D0")),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ])
    story.append(Table(summary_data, colWidths=[W / 4] * 4, style=ts))
    story.append(Spacer(1, 8 * mm))

    if units:
        story.append(Paragraph("Units", H2))
        house_rows = [["Unit", "Monthly Rent (UGX)", "Location", "Current Tenant"]]
        for h, val in sorted(units.items()):
            rent = parse_amount(val["rent"] if isinstance(val, dict) else val)
            location = val.get("location", "—") if isinstance(val, dict) else "—"
            tenant_name = next((t.get("name", "Unnamed") for t in tenants if t.get("unit") == h), "Vacant")
            house_rows.append([h, f"{int(rent):,}", location or "—", tenant_name])
        ht = TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1FAD9F")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F0FAF8")]),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#B0D4D0")),
            ("INNERGRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D0E8E4")),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ])
        story.append(Table(house_rows, colWidths=[W * 0.15, W * 0.22, W * 0.28, W * 0.35], style=ht))
        story.append(Spacer(1, 8 * mm))

    story.append(Paragraph("Tenant Records", H2))
    if not tenants:
        story.append(Paragraph("No tenants on record.", MUTED))
    else:
        for t in tenants:
            history = t.get("payment_history", [])
            dep_history = t.get("deposit_history", [])
            paid_total = int(
                sum(int(r.get("amount", 0)) for r in history if not r.get("_cancelled", False))
                + sum(float(r.get("amount", 0)) for r in dep_history if not r.get("_cancelled", False))
            )
            story.append(Spacer(1, 3 * mm))
            story.append(Paragraph(f"<b>{t.get('name', 'Unnamed Tenant')}</b>", BOLD))
            story.append(Paragraph(
                f"Unit: {t.get('unit', '—')}  |  Rent: UGX {int(parse_amount(t.get('rent', 0))):,}/mo  |  "
                f"{_days_status(t)}", MUTED))
            info_rows = [
                ["Phone", t.get("phone", "—"), "Email", t.get("email", "—")],
                ["Occupation", t.get("occupation", "—"), "", ""],
                ["Entry Date", t.get("entry_date", "—"), "Due Date", t.get("due_date", "—")],
                ["Emergency", t.get("emergency_contact", "—"), "Emg. Phone", t.get("emergency_phone", "—")],
            ]
            it = TableStyle([
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
                ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor("#6E8482")),
                ("TEXTCOLOR", (2, 0), (2, -1), colors.HexColor("#6E8482")),
                ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("#F6FBFA"), colors.white]),
                ("BOX", (0, 0), (-1, -1), 0.3, colors.HexColor("#D0E8E4")),
                ("INNERGRID", (0, 0), (-1, -1), 0.2, colors.HexColor("#E0F0EC")),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ])
            story.append(Spacer(1, 1 * mm))
            story.append(Table(info_rows, colWidths=[W * 0.16, W * 0.34, W * 0.16, W * 0.34], style=it))
            if history:
                story.append(Spacer(1, 2 * mm))
                story.append(Paragraph(f"Payment History  |  Total Income: UGX {paid_total:,}", BOLD))
                pay_rows = [["#", "Date Paid", "Period", "Months", "Amount (UGX)"]]
                for n, rec in enumerate(reversed(history), 1):
                    pay_rows.append([str(n), rec.get("date", "—"),
                                      fmt_period(rec.get("from_date", ""), rec.get("to_date", "")),
                                      str(rec.get("months", 1)), f"{int(rec.get('amount', 0)):,}"])
                pt = TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF7F4")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8),
                    ("ALIGN", (0, 0), (0, -1), "CENTER"),
                    ("ALIGN", (4, 0), (4, -1), "RIGHT"),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F6FBFA")]),
                    ("BOX", (0, 0), (-1, -1), 0.3, colors.HexColor("#B0D4D0")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.2, colors.HexColor("#D8EEEA")),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ])
                story.append(Table(pay_rows, colWidths=[W * 0.06, W * 0.18, W * 0.32, W * 0.12, W * 0.32], style=pt))
            else:
                story.append(Paragraph("No payments recorded yet.", MUTED))
            story.append(HRFlowable(width=W, thickness=0.5, color=colors.HexColor("#C8E0DC"),
                                     spaceAfter=4, spaceBefore=4))
    doc.build(story)
    return PDF_FILE


# ── watchlist / alerts ───────────────────────────────────────────────────
def get_watchlist_tenants(tenants, max_days=10):
    today = date.today()
    items = []
    for t in tenants:
        if status_level(t, today)[0] == "paid":
            continue
        due_str = pending_reference_date_str(t)
        due = _parse_date(due_str)
        if due is None:
            continue
        days_left = (due - today).days
        if days_left <= max_days:
            items.append((t, days_left))
    items.sort(key=lambda pair: pair[1])
    return items

# =========================================================================
#  SECTION 2 -- Flask app + REST API
# =========================================================================
# (app itself is created near the top of the file, before the CLOUD_MODE
# gate's @app.before_request decorator needs it)

# ── device roster + secret key persistence ─────────────────────────────
# Everything about "who is paired" used to live only in this process's
# memory, so restarting the companion (which used to happen on every
# desktop-app "Connect" click) wiped every paired phone's session AND
# forgot every device it had ever seen -- forcing a rescan for a phone
# that had already scanned once before. This file makes both survive a
# restart: the signing secret (so old session cookies keep verifying)
# and the device roster (so a known phone is still known).
DEVICES_FILE = os.path.join(DATA_DIR, "devices.json")
_devices_io_lock = threading.Lock()


def _load_devices_file():
    if not os.path.exists(DEVICES_FILE):
        return {"secret_key": None, "devices": {}}
    try:
        with open(DEVICES_FILE, "r", encoding="utf-8") as f:
            raw = json.load(f)
        if not isinstance(raw, dict):
            raise ValueError("bad devices.json")
        raw.setdefault("secret_key", None)
        raw.setdefault("devices", {})
        return raw
    except Exception:
        return {"secret_key": None, "devices": {}}


def _save_devices_file_locked():
    """Caller must hold _devices_io_lock."""
    tmp = DEVICES_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(_devices_state, f, indent=2)
    os.replace(tmp, DEVICES_FILE)


_devices_state = _load_devices_file()

# A hardcoded fallback secret would be visible to anyone reading this
# file's source (it ships as plain .py, not compiled/obfuscated) --
# which would let them forge a validly-signed session cookie without
# ever entering the PIN or scanning a QR code. So it's still generated
# randomly -- just once, then reused on every later restart so existing
# session cookies and paired devices don't all become invalid just
# because the process restarted. RENTAL_APP_SECRET still overrides it
# for anyone who wants to manage that themselves.
#
# WHERE it's persisted differs by mode:
#   - CLOUD_MODE: in Postgres (cloud_app_secret table), NOT the local
#     devices.json file -- a Render (or similar) web service's local disk
#     is wiped on every restart/redeploy/free-tier spin-down, which used
#     to silently rotate this secret each time, invalidating every
#     phone's pairing cookie and sending them back to the "Waiting to
#     connect" screen for no visible reason.
#   - PC-local mode: devices.json next to the process, as before -- the
#     desktop app's own disk is not ephemeral, so this was never broken
#     there.
if os.environ.get("RENTAL_APP_SECRET"):
    app.secret_key = os.environ["RENTAL_APP_SECRET"]
elif CLOUD_MODE:
    app.secret_key = _cloud_get_or_create_secret()
else:
    if not _devices_state.get("secret_key"):
        _devices_state["secret_key"] = secrets.token_hex(32)
        with _devices_io_lock:
            _save_devices_file_locked()
    app.secret_key = _devices_state["secret_key"]

app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["PERMANENT_SESSION_LIFETIME"] = 60 * 60 * 24 * 180  # 180 days

# Set by the desktop app (via /api/announce-disconnect) right before it
# tears this server down deliberately, e.g. Settings → Disconnect. Lets
# already-connected phones tell that apart from a plain network/PC
# outage: they poll /api/lock-status and see this flag while the server
# is still reachable, instead of only ever finding out via a failed
# request (which is indistinguishable from "PC is off").
_disconnect_state = {"announced": False}

# ── QR-only pairing gate ────────────────────────────────────────────────
# A QR code is nothing more than an encoded URL -- there's no way for a
# server to tell "this request came from a camera scan" apart from "this
# request came from someone pasting/forwarding that same URL", since
# they're literally the same HTTP request. What IS achievable, and is
# what this implements: make the link single-use and short-lived, so a
# forwarded/screenshotted/copied link stops working (serves nothing, not
# even an error that reveals this is a rental-management app) as soon as
# either (a) the original scan succeeds, or (b) the desktop app shows a
# fresh code (e.g. a phone connects, or the rotation timer below ticks).
# The desktop app pushes the currently-valid token here every time it
# displays a QR code (see /api/pairing-token below); the "/" route only
# ever serves the app shell to a request presenting that exact,
# not-yet-consumed token, OR to a request whose X-Device-Id already
# belongs to a known, non-kicked device -- so a phone that scanned once
# never needs to scan again, only a brand-new or disconnected device
# needs a live token.
_pairing_state = {"token": None, "consumed": False}
_pairing_lock = threading.Lock()


def _pairing_ok(device_id=""):
    """True if this request may load the app shell -- see the module
    comment above. Consumes the active token on a successful match so it
    can never work a second time for anyone else it gets forwarded to."""
    if session.get("paired"):
        return True
    if device_id and _device_known(device_id):
        # Recognized via X-Device-Id (e.g. the waiting page's background
        # retry, or an /api/ call) rather than the one-time token. Set the
        # session cookie here too, so the NEXT plain top-level navigation
        # (which can't carry the X-Device-Id header at all) is recognized
        # via session.get("paired") above instead of falling through to
        # the waiting page again.
        session["paired"] = True
        session.permanent = True
        return True
    token = request.args.get("pt", "")
    with _pairing_lock:
        active = _pairing_state["token"]
        matched = bool(active) and bool(token) and token == active and not _pairing_state["consumed"]
        if matched:
            _pairing_state["consumed"] = True
    if matched:
        session["paired"] = True
        session.permanent = True
        if device_id:
            _readmit_device(device_id)
    return matched


def _readmit_device(device_id):
    """Clears a device's 'kicked' flag if -- and only if -- it just
    matched a fresh, valid, single-use pairing token (see the caller
    above): that can only happen because the desktop app explicitly
    showed a new QR code and this device scanned it, i.e. an admin
    re-invited it. Nothing else (a lingering session cookie, a replayed
    /api/ call, the device's own retry loop) can undo a kick this way --
    without this, a disconnected device stayed blocked forever even
    after scanning a brand-new code, since nothing else ever cleared the
    flag. Still respects MAX_DEVICES: if the cap is already full, it
    stays kicked until a slot is freed."""
    import time
    with _devices_lock:
        devices = _devices_dict()
        rec = devices.get(device_id)
        if rec is None or not rec.get("kicked"):
            return
        if sum(1 for r in devices.values() if not r.get("kicked")) >= MAX_DEVICES:
            return
        rec["kicked"] = False
        rec["last_seen"] = time.time()
        _save_devices_locked()

# Tracks which phones/browsers have ever paired with this app, persisted
# to disk (DEVICES_FILE) so the roster survives the companion process
# being stopped and restarted -- e.g. every time the desktop app is
# closed and reopened. Each client generates a random id on first load
# (see DEVICE_ID in the JS below) and sends it as X-Device-Id on every
# /api/ request. Scanning the QR code (or opening an already-known
# device's saved link) grants full access immediately -- there is no
# separate "approve" step; the token gate above is what stands in for
# that. A device stays listed, occupying one of the MAX_DEVICES slots,
# until it is explicitly disconnected from the desktop app -- going
# offline (no internet, phone off, app closed) does NOT drop it from the
# roster, only its displayed status changes.
MAX_DEVICES = 3           # Settings -> Connect Phone allows this many phones
                          # connected at once. A 4th scan sees a "device
                          # limit reached" screen; disconnect one from the
                          # desktop app's device list to free a slot.
ONLINE_TIMEOUT = 20       # seconds since last contact before a device's
                          # displayed status flips from "online" to
                          # "offline" -- display only, never removes it.
_devices_lock = threading.Lock()


def _devices_dict():
    """The live device roster, backed by _devices_state (loaded from /
    persisted to DEVICES_FILE). Caller must hold _devices_lock for any
    read that needs to stay consistent with a following write."""
    return _devices_state["devices"]


def _save_devices_locked():
    """Persist the current roster to disk. Caller must hold _devices_lock."""
    with _devices_io_lock:
        _save_devices_file_locked()


def _device_known(device_id):
    """Non-mutating: has this device ever paired and not since been
    disconnected? Used by _pairing_ok so a previously-scanned phone can
    reload the app shell without a fresh token."""
    if not device_id:
        return False
    with _devices_lock:
        rec = _devices_dict().get(device_id)
        return bool(rec and not rec.get("kicked"))


def _device_was_kicked(device_id):
    """Non-mutating: has this device been explicitly disconnected from
    the desktop app's device list?"""
    if not device_id:
        return False
    with _devices_lock:
        rec = _devices_dict().get(device_id)
        return bool(rec and rec.get("kicked"))


# Best-effort map of iPhone physical screen resolution ("WIDTHxHEIGHT")
# to model name(s). Apple's mobile Safari deliberately doesn't expose the
# exact hardware model in its User-Agent (unlike Android), so this is the
# closest available signal -- and it's inherently approximate, since
# several generations share an identical screen. Where more than one
# model shares a resolution, all of them are listed.
_IPHONE_SCREEN_MODELS = {
    "1170x2532": "iPhone 12/12 Pro/13/13 Pro/14",
    "1284x2778": "iPhone 12 Pro Max/13 Pro Max/14 Plus",
    "1179x2556": "iPhone 14 Pro/15/15 Pro",
    "1290x2796": "iPhone 14 Pro Max/15 Plus/15 Pro Max",
    "1206x2622": "iPhone 16/16 Pro",
    "1320x2868": "iPhone 16 Plus/16 Pro Max",
    "1080x2340": "iPhone 12 mini/13 mini",
    "828x1792": "iPhone 11/XR",
    "1125x2436": "iPhone X/XS/11 Pro",
    "1242x2688": "iPhone XS Max/11 Pro Max",
    "750x1334": "iPhone 6/6s/7/8/SE (2nd/3rd gen)",
    "1080x1920": "iPhone 6/7/8 Plus",
    "640x1136": "iPhone 5/5s/5c/SE (1st gen)",
}


def _guess_ios_model(is_ipad, screen_hint):
    """`screen_hint` is the "WIDTHxHEIGHT@DPR" string sent by the client
    (see DEVICE_SCREEN_HINT in the JS below). Converts logical CSS pixels
    back to physical pixels (x DPR) and looks that up. Falls back to the
    generic platform name if the resolution isn't recognized or wasn't
    sent."""
    if is_ipad:
        return "iPad"  # too many iPad models share resolutions to guess further
    try:
        dims, dpr = screen_hint.split("@")
        w, h = (int(n) for n in dims.split("x"))
        dpr = float(dpr)
        w, h = round(w * dpr), round(h * dpr)
        key = f"{min(w,h)}x{max(w,h)}"
        return _IPHONE_SCREEN_MODELS.get(key, "iPhone")
    except Exception:
        return "iPhone"


def _label_for_user_agent(ua, screen_hint=""):
    """Turns a raw User-Agent (plus an optional screen-resolution hint)
    into a short, phone-model-only label for the admin device list --
    e.g. 'iPhone 12', 'Tecno KI5k', 'itel A662L'. Browser name is
    intentionally never included (admin request: only the device/phone
    type should be shown, not which browser it's using)."""
    ua = ua or ""
    if "iPhone" in ua or "iPad" in ua:
        return _guess_ios_model("iPad" in ua, screen_hint)
    if "Android" in ua:
        # Standard Android UA shape: "...; Android <ver>; <MODEL>) ...",
        # sometimes followed by "Build/...". The model segment is where
        # manufacturers put their real marketing/model name (e.g. "Tecno
        # KI5k", "itel A662L", "SM-A125F", "Redmi Note 11") -- this is
        # the one platform where the exact device name really is
        # available, so use it verbatim instead of just "Android".
        m = re.search(r"Android\s+[\d.]+\s*;\s*([^;)]+?)(?:\s+Build/[^;)]+)?\)", ua)
        if m:
            model = m.group(1).strip()
            if model and model.lower() not in ("k", ""):
                return model
        return "Android"
    if "Windows" in ua:
        return "Windows PC"
    if "Macintosh" in ua:
        return "Mac"
    return "Device"


def _touch_device(device_id, user_agent=None, screen_hint="", fingerprint=""):
    """Refresh an already-known device's last-seen time, or register a
    brand-new one -- immediately admitted, no separate approval step --
    if there's room under MAX_DEVICES. Returns (admitted, canonical_id):
    admitted is False only if it was turned away because the cap is
    already full or it was explicitly disconnected from the desktop
    app's device list. Existing devices (online or not) are never
    evicted to make room for a newcomer -- a slot only frees up when the
    desktop app disconnects one.

    Same fingerprint-based reconciliation as the cloud counterpart (see
    _cloud_touch_device): an unrecognized device_id whose fingerprint
    matches an already-known device reuses that device's slot instead of
    registering a new one, and canonical_id tells the caller which id
    that was."""
    if not device_id:
        return True, device_id
    import time
    with _devices_lock:
        devices = _devices_dict()
        rec = devices.get(device_id)
        if rec is not None and rec.get("kicked"):
            return False, device_id
        now = time.time()
        if rec is not None:
            rec["last_seen"] = now
            # Re-derive the label on every poll rather than only at first
            # registration: the very first request from a device can
            # arrive before its screen-hint header is populated in some
            # edge cases, so this lets the label upgrade from a generic
            # "iPhone"/"Android" to a specific model as soon as a usable
            # hint shows up. Skipped entirely once the person has given
            # this device a custom name (see _set_local_device_label) --
            # that name should never be silently overwritten again.
            if not rec.get("custom_label_locked"):
                fresh_label = _label_for_user_agent(user_agent, screen_hint)
                if fresh_label and fresh_label not in ("Device",):
                    rec["label"] = fresh_label
            if fingerprint:
                rec["fingerprint"] = fingerprint
            _save_devices_locked()
            return True, device_id
        if fingerprint:
            match_id = next(
                (d for d, r in devices.items()
                 if not r.get("kicked") and r.get("fingerprint") == fingerprint),
                None)
            if match_id:
                match_rec = devices[match_id]
                match_rec["last_seen"] = now
                if not match_rec.get("custom_label_locked"):
                    fresh_label = _label_for_user_agent(user_agent, screen_hint)
                    if fresh_label and fresh_label not in ("Device",):
                        match_rec["label"] = fresh_label
                _save_devices_locked()
                return True, match_id
        if sum(1 for r in devices.values() if not r.get("kicked")) >= MAX_DEVICES:
            return False, device_id
        devices[device_id] = {
            "first_seen": now,
            "last_seen": now,
            "label": _label_for_user_agent(user_agent, screen_hint),
            "kicked": False,
            "fingerprint": fingerprint,
        }
        _save_devices_locked()
        return True, device_id


def _device_admitted(device_id):
    """Non-mutating check: is this device allowed to use every other
    /api/ call? True for any known, non-kicked device -- pairing (scan
    or already-known link) is the only gate; there is no further
    approval step."""
    if not device_id:
        return True
    with _devices_lock:
        rec = _devices_dict().get(device_id)
        return bool(rec and not rec.get("kicked"))


def _active_device_count():
    """How many devices are currently on the roster (online or not) --
    what the Connect Phone card's 'Connected -- N/3' status and the
    MAX_DEVICES cap are based on. A device only stops counting once it
    is explicitly disconnected."""
    with _devices_lock:
        return sum(1 for rec in _devices_dict().values() if not rec.get("kicked"))


def _list_devices():
    """Snapshot of every currently-tracked (non-kicked) device for the
    desktop app's admin list: id (full, used to target a kick), a short
    id for display, a human label, whether it's currently reachable
    (online), and when it last synced (epoch seconds) -- devices stay
    listed indefinitely once offline, they are never dropped by a
    timeout, only by an explicit Disconnect."""
    import time
    with _devices_lock:
        now = time.time()
        return [
            {
                "device_id": d,
                "short_id": d[:8],
                "label": rec.get("label") or "Unknown device",
                "last_seen": rec["last_seen"],
                "online": (now - rec["last_seen"]) <= ONLINE_TIMEOUT,
            }
            for d, rec in sorted(_devices_dict().items(),
                                  key=lambda kv: kv[1]["last_seen"], reverse=True)
            if not rec.get("kicked")
        ]


def _kick_device(device_id):
    """Forcibly disconnects one specific phone/browser (desktop app's
    Connect Phone -> device list -> Disconnect), freeing its slot.
    Marked kicked (rather than deleted outright) so it's permanently
    blocked from silently re-admitting itself on its next poll -- the
    next request from that device_id gets turned away instead. It clears
    only if the phone/browser generates a fresh device id (e.g. clearing
    site data) or is manually removed from devices.json."""
    with _devices_lock:
        rec = _devices_dict().get(device_id)
        if rec is not None:
            rec["kicked"] = True
        else:
            _devices_dict()[device_id] = {
                "first_seen": 0, "last_seen": 0,
                "label": "Unknown device", "kicked": True,
            }
        _save_devices_locked()


def _set_local_device_label(device_id, label):
    """Lets a phone give itself a custom name (e.g. 'Mary's iPhone')
    instead of the generic auto-detected model name -- and locks it so
    _touch_device never overwrites it again on a later poll."""
    import time
    with _devices_lock:
        devices = _devices_dict()
        rec = devices.get(device_id)
        now = time.time()
        if rec is None:
            rec = {"first_seen": now, "last_seen": now, "kicked": False, "fingerprint": ""}
            devices[device_id] = rec
        rec["label"] = label
        rec["custom_label_locked"] = True
        rec["last_seen"] = now
        _save_devices_locked()


# ── auth helpers ─────────────────────────────────────────────────────────
def _pin_required(data):
    return bool(data.get("settings", {}).get("pin_hash"))


def _authed():
    return session.get("unlocked") is True


@app.before_request
def _guard():
    if CLOUD_MODE:
        return None
    if request.path.startswith("/api/"):
        if request.path == "/api/lock-status":
            return None
        if (request.path == "/api/devices" or request.path.startswith("/api/devices/")
                or request.path in ("/api/shutdown", "/api/announce-disconnect",
                                     "/api/device-count", "/api/pairing-token")) and (
                request.remote_addr in ("127.0.0.1", "::1")
                or (request.remote_addr or "").startswith("::ffff:127.")):
            # Only the desktop app (calling from the same PC) may stop the
            # server, announce it's about to, list/kick devices, or push a
            # fresh pairing token, this way; a phone reaching over the LAN
            # still needs to unlock first like any other /api/ call.
            return None
        device_id = request.headers.get("X-Device-Id", "")
        if device_id and _device_was_kicked(device_id):
            return jsonify({"error": "kicked"}), 403
        if device_id and not _device_admitted(device_id):
            return jsonify({"error": "device_limit_reached",
                             "max_devices": MAX_DEVICES}), 403
        if request.path == "/api/unlock":
            return None
        data = load_state()
        if _pin_required(data) and not _authed():
            return jsonify({"error": "locked"}), 401
    return None


# ── page ─────────────────────────────────────────────────────────────────
WAITING_HTML = """<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1, viewport-fit=cover">
<title>Tenant Management</title>
<meta name="theme-color" content="#0E4F4F">
<link rel="icon" href="icon-192.png">
<link rel="apple-touch-icon" href="icon-256.png">
<style>
  html,body{margin:0;padding:0;background:#0E4F4F;color:#fff;font-family:-apple-system,'Segoe UI',Inter,sans-serif;}
  .wrap{min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;padding:24px;}
  .mark{font-size:40px;margin-bottom:10px;}
  h1{font-size:19px;margin:0 0 8px;}
  p{font-size:13px;opacity:.8;max-width:280px;margin:0;}
  .spin{margin-top:22px;width:22px;height:22px;border-radius:50%;border:3px solid rgba(255,255,255,.25);
        border-top-color:#fff;animation:sp 0.9s linear infinite;}
  @keyframes sp{to{transform:rotate(360deg);}}
</style></head>
<body><div class="wrap">
  <div class="mark">🏠</div>
  <h1>Tenant Management</h1>
  <p>Waiting to connect. Scan the QR code shown in Connect Phone on the PC to open this device's tenant data.</p>
  <div class="spin"></div>
</div>
<script>
  // Not paired yet (or a forwarded/expired link) -- reveals nothing about
  // tenant data. Quietly retries in the background so this page moves on
  // by itself the moment the desktop app admits this device, without the
  // person having to notice and reload manually.
  setInterval(function () {
    fetch(location.pathname + location.search, {headers:{'X-Device-Id': (function(){
      try { return localStorage.getItem('rm_device_id_v1') || ''; } catch(e) { return ''; }
    })()}}).then(function (r) { if (r.ok) location.reload(); }).catch(function () {});
  }, 4000);
</script>
</body></html>"""


# ── page ─────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    if CLOUD_MODE:
        # sid+key identify WHICH household's data this is and are what
        # every subsequent /api/ call authenticates with going forward --
        # they can't be single-use themselves, or the phone couldn't keep
        # syncing after this first load. Getting INTO the app the first
        # time is gated separately, by pt: the one-time token the desktop
        # pushes every time it (re)shows the QR code (see
        # _cloud_pairing_ok() above). A link with valid sid+key but no
        # valid pt -- forwarded, screenshotted, or just reopened by
        # someone who was never shown the QR -- gets the same branded
        # waiting page a stale LAN link gets, never the app shell.
        session_id = request.args.get("sid", "")
        secret_key = request.args.get("key", "")
        pt = request.args.get("pt", "")
        row = _cloud_get_row(session_id) if session_id else None
        if not session_id or not secret_key or not row or row["secret_key"] != secret_key:
            return Response("", status=404, mimetype="text/plain")
        if not _cloud_pairing_ok(session_id, pt):
            return Response(WAITING_HTML, status=404, mimetype="text/html")
        bootstrap = (
            "<script>window.__CLOUD_DIRECT__=" + json.dumps({
                "sessionId": session_id, "secretKey": secret_key,
            }) + ";</script>\n"
        )
        html = INDEX_HTML.replace("<head>", "<head>\n" + bootstrap, 1)
        # Carry sid/key onto the manifest request too. manifest.json's
        # start_url is resolved against the MANIFEST's own URL, not this
        # page's -- so without this, an installed home-screen icon always
        # launches a bare "/" with no sid/key at all. index() then 404s
        # with an EMPTY body (see the sid/key check above), which is
        # exactly the blank white page on launch. Passing them through
        # here lets manifest() below bake them into start_url itself.
        qs = f"sid={quote(session_id)}&key={quote(secret_key)}"
        html = html.replace(
            'href="manifest.json"', f'href="manifest.json?{qs}"', 1)
        return Response(html, mimetype="text/html")
    if not _pairing_ok(request.headers.get("X-Device-Id", "")):
        # No token, a stale/wrong one, or one already used once -- this
        # used to be a totally bare, unlabelled 404 so a forwarded/copied
        # link couldn't even tell this was a rental-management app. That
        # also meant the browser tab (and "Add to Home Screen") had no
        # <title> to show and fell back to displaying the raw URL
        # instead. This branded waiting page keeps the same "no tenant
        # data leaks to an unpaired visitor" guarantee, but shows the
        # Tenant Management name/icon instead of the bare link, and
        # quietly retries in the background until it's admitted.
        return Response(WAITING_HTML, status=404, mimetype="text/html")
    return Response(INDEX_HTML, mimetype="text/html")


# ── PWA: manifest, service worker, icon ───────────────────────────────────
# These three routes are what let a phone or desktop browser "install" this
# page (Chrome/Edge: address-bar install icon or ⋮ → Install app; Safari:
# Share → Add to Home Screen) so it opens full-screen with its own icon,
# just like a native app — no separate hosting or build step required,
# since it's all served straight from this same process.
@app.route("/icon-<int:size>.png")
def app_icon(size):
    icon_bytes = base64.b64decode(APP_ICON_256_B64)
    return Response(icon_bytes, mimetype="image/png")


@app.route("/manifest.json")
def manifest():
    # In CLOUD_MODE, index() now forwards this page's sid/key onto the
    # manifest request (see its "?sid=&key=" href rewrite above) so they
    # can be baked into start_url here. Without this, start_url resolves
    # to a bare "/" -- valid per spec, since these fields resolve against
    # the MANIFEST's own URL rather than the page's -- and a home-screen
    # launch of that bare "/" gets index()'s empty-body 404 (no session
    # means no household to show), i.e. the reported blank white page.
    sid = request.args.get("sid", "")
    key = request.args.get("key", "")
    start_url = f"./?sid={quote(sid)}&key={quote(key)}" if (CLOUD_MODE and sid and key) else "."
    return jsonify({
        "name": "Tenant Management",
        "short_name": "Tenant Management",
        "description": "Manage tenants, payments, and units on the go.",
        # Relative, not "/": per the manifest spec these all resolve
        # against the manifest's OWN url, not the page's.
        "start_url": start_url,
        "scope": ".",
        "display": "standalone",
        "orientation": "portrait",
        "background_color": "#0E4F4F",
        "theme_color": "#0E4F4F",
        "icons": [
            {"src": "icon-192.png", "sizes": "192x192", "type": "image/png", "purpose": "any maskable"},
            {"src": "icon-256.png", "sizes": "256x256", "type": "image/png", "purpose": "any maskable"},
        ],
    })


@app.route("/sw.js")
def service_worker():
    # Cache-first for the app shell (so it still opens with no signal),
    # but always goes to the network for /api/* so tenant data is never
    # served stale. Scope is "/" via the header below so it can control
    # the whole app, not just its own folder.
    js = """
// Bumped so every existing install picks up this fix on its next
// activate() instead of continuing to serve a stale precached "./" that
// (pre-fix) never had sid/key on it in the first place.
const CACHE = 'rental-app-shell-v5';
// sid/key travel here via this script's OWN url (see the
// navigator.serviceWorker.register('sw.js?sid=...&key=...') call) --
// self.location.search is that query string, readable from inside the
// worker regardless of which page registered it.
const PARAMS = self.location.search;
const ROOT_URL = './' + PARAMS;
const MANIFEST_URL = './manifest.json' + PARAMS;
// Relative, not root-absolute: resolved against this script's own URL.
const SHELL_URLS = [ROOT_URL, MANIFEST_URL, './icon-192.png'];

self.addEventListener('install', (evt) => {
  self.skipWaiting();
  evt.waitUntil(caches.open(CACHE).then((c) => c.addAll(SHELL_URLS)));
});

self.addEventListener('activate', (evt) => {
  evt.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(keys.filter((k) => k !== CACHE).map((k) => caches.delete(k)))
    ).then(() => self.clients.claim())
  );
});

self.addEventListener('fetch', (evt) => {
  const url = new URL(evt.request.url);
  if (url.pathname.includes('/api/')) return;

  // Network-first for the app shell: whenever the phone can reach the
  // server, it always gets whatever HTML/JS is currently running there,
  // so edits to this app show up the next time it's opened. Falls back
  // to cache whenever the network request fails outright (no Wi-Fi/data,
  // or the browser was closed and reopened somewhere without a
  // connection) so the app keeps opening instead of showing nothing --
  // and if even THIS exact request was never cached before, falls back
  // to ROOT_URL (the paired shell, WITH sid/key) rather than the bare
  // worker scope, which used to have no session on it at all and is
  // what produced the blank white page after installing to the home
  // screen.
  if (evt.request.method === 'GET') {
    evt.respondWith(
      fetch(evt.request).then((resp) => {
        if (resp && resp.ok) {
          caches.open(CACHE).then((c) => c.put(evt.request, resp.clone()));
          return resp;
        }
        // Not a real page -- our own not-yet-paired waiting page, or a
        // genuine server error. Prefer the last cached shell if there is
        // one, so the app still opens.
        return caches.match(evt.request).then((cached) => cached || caches.match(ROOT_URL) || resp);
      }).catch(() => caches.match(evt.request).then((cached) => cached || caches.match(ROOT_URL)))
    );
  }
});
"""
    return Response(js, mimetype="application/javascript",
                     headers={"Service-Worker-Allowed": "/"})


# ── cloud durability / sync ─────────────────────────────────────────────
CLOUD_SYNC_FILE = os.path.join(DATA_DIR, "cloud_sync.json")


def _load_cloud_sync_config():
    """Local-only file the desktop app owns (creates/updates), telling the
    phone where the cloud fallback service lives and how to authenticate
    to it. Read-only from here -- if it doesn't exist yet, cloud fallback
    just isn't configured for this install."""
    if not os.path.exists(CLOUD_SYNC_FILE):
        return None
    try:
        with open(CLOUD_SYNC_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        if cfg.get("cloud_base_url") and cfg.get("session_id") and cfg.get("secret_key"):
            return cfg
    except Exception:
        pass
    return None


@app.route("/api/cloud-config")
def cloud_config():
    """PC-mode only: tells the phone (while it can still reach the PC)
    where the cloud fallback is and how to authenticate to it, so the
    phone can keep working -- reading AND writing -- once the PC goes
    offline. The phone's own api() wrapper caches this response like any
    other GET, so it's still available from the phone's local cache even
    after the PC is gone."""
    cfg = _load_cloud_sync_config()
    if not cfg:
        return jsonify({"configured": False})
    return jsonify({
        "configured": True,
        "cloud_base_url": cfg["cloud_base_url"].rstrip("/"),
        "session_id": cfg["session_id"],
        "secret_key": cfg["secret_key"],
    })


@app.route("/api/_sync", methods=["GET", "PUT"])
def cloud_sync():
    """Cloud-mode only (see _cloud_gate). GET returns the stored snapshot
    for reconciliation; PUT accepts a full snapshot push -- from either
    the PC (after a local save) or, in principle, anywhere else -- and
    merges it record-by-record against whatever's already stored (see
    _merge_cloud_data) rather than treating either side's whole snapshot
    as "the winner"."""
    if not CLOUD_MODE:
        return Response("", status=404, mimetype="text/plain")

    if request.method == "GET":
        row = _cloud_get_row(g.session_id)
        if not row:
            return jsonify({"exists": False})
        return jsonify({
            "exists": True,
            "data": row["data"],
            "updated_at": row["updated_at"].isoformat(),
            "updated_by": row["updated_by"],
        })

    body = request.get_json(silent=True) or {}
    incoming_data = body.get("data")
    updated_by = body.get("updated_by", "unknown")
    if not isinstance(incoming_data, dict):
        return jsonify({"ok": False, "error": "bad_request"}), 400

    # Merged record-by-record against the currently stored snapshot,
    # rather than being accepted or rejected wholesale based on comparing
    # the PUSHING side's own wall-clock timestamp (body.get("updated_at"))
    # against the row's stored timestamp. That comparison used to be the
    # gate here, but every ACCEPTED write is stamped with the database's
    # own now() (see _cloud_save) -- an authoritative, ever-advancing
    # clock -- while the pushing side (e.g. the PC, syncing every ~20s)
    # sends its own local clock's reading. Even a few seconds of
    # client/server clock drift meant that, after the first successful
    # push, every later push from that same side looked "not new enough"
    # and was silently discarded forever -- while writes landing directly
    # on this server (from the phone/web app) always kept advancing the
    # row's timestamp, so pulls from that side (reconcile) kept working.
    # That's what made syncing look one-directional: phone → PC worked,
    # PC → phone silently stopped. Merging here removes the clock
    # comparison from the equation entirely; per-record `_updated_at`
    # stamps (see _stamp_changed_records) already make the merge safe
    # regardless of which side's clock is ahead.
    row = _cloud_get_row(g.session_id)
    merged = _merge_cloud_data(incoming_data, row["data"]) if row else incoming_data
    server_updated_at = save_raw(merged, updated_by=updated_by)
    return jsonify({"ok": True, "stored": True, "updated_at": server_updated_at, "data": merged})



@app.route("/api/lock-status")
def lock_status():
    if CLOUD_MODE:
        # No PC in the loop here at all -- this only ever confirms the
        # cloud service itself is reachable (which is what a cloud-direct
        # phone's isOnline badge actually means: "works with the PC off"),
        # plus this device's own kicked/roster status. _cloud_gate already
        # ran the touch/kicked check above before this handler is reached.
        device_id = request.headers.get("X-Device-Id", "")
        return jsonify({
            "pin_set": False,
            "unlocked": True,
            "app_name": APP_NAME,
            "disconnecting": False,
            "kicked": _cloud_device_kicked(g.session_id, device_id),
            "device_count": _cloud_active_device_count(g.session_id),
            "max_devices": MAX_DEVICES,
            "device_limit_reached": False,
            "data_updated_at": _cloud_get_updated_at(g.session_id),
        })
    data = load_state()
    device_id = request.headers.get("X-Device-Id", "")
    kicked = _device_was_kicked(device_id)
    admitted = True
    canonical_device_id = None
    if not _disconnect_state["announced"] and not kicked:
        admitted, canonical_device_id = _touch_device(
            device_id,
            request.headers.get("User-Agent"),
            request.headers.get("X-Device-Screen", ""),
            request.headers.get("X-Device-Fingerprint", ""),
        )
    try:
        data_updated_at = str(os.path.getmtime(DATA_FILE))
    except OSError:
        data_updated_at = ""
    resp = {
        "pin_set": _pin_required(data),
        "unlocked": _authed(),
        "app_name": APP_NAME,
        "disconnecting": _disconnect_state["announced"],
        "kicked": kicked,
        "device_count": _active_device_count(),
        "max_devices": MAX_DEVICES,
        "device_limit_reached": (not admitted) and not kicked,
        "data_updated_at": data_updated_at,
    }
    if canonical_device_id and canonical_device_id != device_id:
        resp["canonical_device_id"] = canonical_device_id
    return jsonify(resp)


@app.route("/api/device-count")
def device_count():
    """Polled locally (127.0.0.1) by the desktop app's Connect Phone panel
    to tell Pending (server up, 0 phones actively polling) apart from
    Connected (>=1 phone actively polling), and to show how many out of
    the MAX_DEVICES cap are currently in use."""
    return jsonify({"device_count": _active_device_count(), "max_devices": MAX_DEVICES})


@app.route("/api/devices")
def devices_list():
    """The desktop app's Connect Phone card admin list -- which phones/
    browsers are currently connected, with enough of an id shown
    (short_id) and a friendly label to tell them apart, plus a device_id
    to target with /api/devices/<id>/kick. In CLOUD_MODE this is
    session-authenticated (X-Session-Id/X-Secret-Key) rather than
    loopback-only, since the desktop app reaches it over the same public
    URL every phone does -- there's no "same machine" signal to check
    instead, the household's secret_key IS the admin credential here,
    same as it already is for every other cloud endpoint."""
    if CLOUD_MODE:
        return jsonify({"devices": _cloud_list_devices(g.session_id)})
    return jsonify({"devices": _list_devices()})


@app.route("/api/devices/<device_id>/kick", methods=["POST"])
def devices_kick(device_id):
    """Forcibly disconnects one specific phone/browser -- the admin
    'Disconnect' action next to a device in the desktop app's Connect
    Phone card, as opposed to the single Disconnect button that tears
    the whole companion down for every device. See devices_list() above
    re: CLOUD_MODE auth."""
    if CLOUD_MODE:
        _cloud_kick_device(g.session_id, device_id)
    else:
        _kick_device(device_id)
    return jsonify({"ok": True})


@app.route("/api/devices/label", methods=["POST"])
def set_device_label():
    """Lets the calling device (identified by its own X-Device-Id header)
    give itself a custom name -- e.g. 'Mary's iPhone' instead of the
    auto-detected 'iPhone 12' -- so it's recognizable in the Connected
    Devices list. Persisted server-side (and locked against being
    overwritten by auto-detection again), so it only needs to be typed
    once, ever, from any device that's already paired -- not re-entered
    each time the web app is opened."""
    device_id = request.headers.get("X-Device-Id", "")
    body = request.get_json(force=True) or {}
    label = (body.get("label") or "").strip()[:40]
    if not device_id:
        return jsonify({"ok": False, "error": "Missing device id."}), 400
    if not label:
        return jsonify({"ok": False, "error": "Please enter a name."}), 400
    if CLOUD_MODE:
        _cloud_set_device_label(g.session_id, device_id, label)
    else:
        _set_local_device_label(device_id, label)
    return jsonify({"ok": True, "label": label})


@app.route("/api/pairing-token", methods=["POST"])
def set_pairing_token():
    """The desktop app calls this every time it displays a QR code,
    handing over the one token that will be accepted by '/' -- see
    _pairing_ok() / _cloud_pairing_ok() above. Immediately invalidates
    whatever token was active before, so an old screenshot or a link
    copied from an earlier QR code stops working the moment a new one is
    shown, not just when it's actually used."""
    body = request.get_json(force=True) or {}
    token = (body.get("token") or "").strip()
    if CLOUD_MODE:
        _cloud_set_pairing_token(g.session_id, token or None)
        return jsonify({"ok": True})
    with _pairing_lock:
        _pairing_state["token"] = token or None
        _pairing_state["consumed"] = False
    return jsonify({"ok": True})


@app.route("/api/announce-disconnect", methods=["POST"])
def announce_disconnect():
    """Called by the desktop app right before it stops this server via
    Settings → Disconnect. Marks the flag picked up by lock_status()
    above so connected phones can show a proper 'Disconnected' state
    instead of just going offline."""
    _disconnect_state["announced"] = True
    return jsonify({"ok": True})


@app.route("/api/unlock", methods=["POST"])
def unlock():
    body = request.get_json(force=True) or {}
    pin = (body.get("pin") or "").strip()
    data = load_state()
    pin_hash = data.get("settings", {}).get("pin_hash", "")
    if not pin_hash:
        session["unlocked"] = True
        return jsonify({"ok": True})
    if hash_secret(pin) == pin_hash:
        session["unlocked"] = True
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "Incorrect PIN."}), 400


@app.route("/api/lock", methods=["POST"])
def lock():
    session["unlocked"] = False
    return jsonify({"ok": True})


@app.route("/api/settings/pin", methods=["POST"])
def set_pin():
    body = request.get_json(force=True) or {}
    new_pin = (body.get("new_pin") or "").strip()
    current_pin = (body.get("current_pin") or "").strip()
    data = load_state()
    settings = data.setdefault("settings", {})
    existing_hash = settings.get("pin_hash", "")

    if existing_hash and hash_secret(current_pin) != existing_hash:
        return jsonify({"ok": False, "error": "Current PIN is incorrect."}), 400
    if not new_pin or not new_pin.isdigit() or len(new_pin) < 4:
        return jsonify({"ok": False, "error": "PIN must be at least 4 digits."}), 400

    settings["pin_hash"] = hash_secret(new_pin)
    settings["lock_mode"] = "pin"
    save_state(data)
    session["unlocked"] = True
    return jsonify({"ok": True})


@app.route("/api/settings/pin", methods=["DELETE"])
def remove_pin():
    body = request.get_json(force=True) or {}
    current_pin = (body.get("current_pin") or "").strip()
    data = load_state()
    settings = data.setdefault("settings", {})
    existing_hash = settings.get("pin_hash", "")
    if existing_hash and hash_secret(current_pin) != existing_hash:
        return jsonify({"ok": False, "error": "Current PIN is incorrect."}), 400
    settings.pop("pin_hash", None)
    settings.pop("lock_mode", None)
    save_state(data)
    return jsonify({"ok": True})


# ── dashboard ────────────────────────────────────────────────────────────
@app.route("/api/dashboard")
def dashboard():
    data = load_state()
    tenants = _with_index(data["tenants"])
    units = data["units"]
    today = date.today()
    month_name = today.strftime("%B")
    this_month = today.strftime("%Y-%m")

    total_collected = int(
        sum(int(r.get("amount", 0)) for t in tenants for r in t.get("payment_history", [])
            if not r.get("_cancelled", False))
        + sum(float(r.get("amount", 0)) for t in tenants for r in t.get("deposit_history", [])
              if not r.get("_cancelled", False))
    )
    counts = {"paid": 0, "underpaid": 0, "pending": 0}
    for t in tenants:
        level, _ = status_level(t, today)
        counts[level] += 1

    occupied_units = {t.get("unit") for t in tenants if t.get("unit")}
    vacant = [u for u in units if u not in occupied_units]
    occupied = len(units) - len(vacant)

    # This-calendar-month income breakdown — mirrors the desktop app's
    # Dashboard "TOTAL INCOME" card exactly (full payments + deposits,
    # minus anything cancelled, all restricted to the current month).
    full_payment_total = int(
        sum(int(r.get("amount", 0)) for t in tenants for r in t.get("payment_history", [])
            if r.get("date", "").startswith(this_month) and not r.get("_cancelled", False)))
    deposit_total = int(
        sum(float(r.get("amount", 0)) for t in tenants for r in t.get("deposit_history", [])
            if r.get("date", "").startswith(this_month) and not r.get("_cancelled", False)))
    cancelled_total = int(
        sum(int(r.get("amount", 0)) for t in tenants for r in t.get("payment_history", [])
            if r.get("date", "").startswith(this_month) and r.get("_cancelled", False))
        + sum(float(r.get("amount", 0)) for t in tenants for r in t.get("deposit_history", [])
              if r.get("date", "").startswith(this_month) and r.get("_cancelled", False)))
    month_income = full_payment_total + deposit_total

    watchlist = get_watchlist_tenants(tenants, max_days=10)
    watchlist_out = [_tenant_summary(t, today) | {"days_left": d} for t, d in watchlist[:8]]

    return jsonify({
        "app_name": APP_NAME,
        "total_tenants": len(tenants),
        "total_units": len(units),
        "vacant_units": len(vacant),
        "occupied_units": occupied,
        "total_collected": total_collected,
        "counts": counts,
        "month_name": month_name,
        "month_income": month_income,
        "full_payment_total": full_payment_total,
        "deposit_total": deposit_total,
        "cancelled_total": cancelled_total,
        "watchlist": watchlist_out,
    })


def _tenant_summary(t, today):
    level, label = status_level(t, today)
    due = _parse_date(t.get("due_date", "")) or _parse_date(t.get("entry_date", ""))
    days_left = (due - today).days if due else None
    dep_paid, dep_remaining, dep_cleared, dep_in_progress = current_deposit_cycle(t)
    return {
        "index": t.get("_idx"),
        "name": t.get("name", "Unnamed"),
        "unit": t.get("unit", "—"),
        "phone": t.get("phone", ""),
        "rent": parse_amount(t.get("rent", 0)),
        "status": t.get("status", "Pending"),
        "level": level,
        "label": label,
        "due_date": t.get("due_date", ""),
        "entry_date": t.get("entry_date", ""),
        "days_left": days_left,
        "deposit_paid": dep_paid,
        "deposit_remaining": dep_remaining,
        "rent_increase_due": rent_increase_due(t),
    }


def _with_index(tenants):
    for i, t in enumerate(tenants):
        t["_idx"] = i
    return tenants


# ── tenants ──────────────────────────────────────────────────────────────
@app.route("/api/tenants")
def list_tenants():
    data = load_state()
    tenants = _with_index(data["tenants"])
    today = date.today()
    q = (request.args.get("q") or "").strip().lower()
    flt = (request.args.get("filter") or "all").lower()

    out = []
    for t in tenants:
        summary = _tenant_summary(t, today)
        if q and q not in summary["name"].lower() and q not in summary["unit"].lower():
            continue
        if flt != "all" and summary["level"] != flt:
            continue
        out.append(summary)
    out.sort(key=lambda s: (s["days_left"] if s["days_left"] is not None else 99999))
    return jsonify({"tenants": out})


@app.route("/api/tenants/<int:idx>")
def get_tenant(idx):
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    t = tenants[idx]
    today = date.today()
    detail = _tenant_summary(t, today) | {
        "index": idx,
        "email": t.get("email", ""),
        "occupation": t.get("occupation", ""),
        "emergency_contact": t.get("emergency_contact", ""),
        "emergency_phone": t.get("emergency_phone", ""),
        "notes": t.get("notes", ""),
        "payment_history": list(reversed(t.get("payment_history", []))),
        "deposit_history": list(reversed(t.get("deposit_history", []))),
        "arrears_history": list(reversed(t.get("arrears_history", []))),
    }
    return jsonify({"tenant": detail})


@app.route("/api/tenants/<int:idx>/months")
def get_tenant_months(idx):
    """Feeds the month-picker dropdown: the already-paid months (from the
    tenant's move-in month up to their current due date -- shown locked,
    can't be re-selected) and the still-open months going forward (shown
    as tickable checkboxes, must be ticked in order starting from the
    first one)."""
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    t = tenants[idx]
    try:
        horizon = min(60, max(1, int(request.args.get("horizon", MONTH_PICKER_HORIZON))))
    except (ValueError, TypeError):
        horizon = MONTH_PICKER_HORIZON
    return jsonify({
        "cleared": cleared_months_list(t),
        "open": open_months_list(t, horizon),
    })


@app.route("/api/history")
def history():
    """Mirrors the desktop app's Records tab: one card per tenant (Entry
    Date + a merged, date-descending ledger of their payment_history and
    deposit_history — cancelled records flagged in place, exactly as on
    desktop), filtered by the same name/unit/phone search."""
    data = load_state()
    tenants = _with_index(data["tenants"])
    q = (request.args.get("q") or "").strip().lower()

    out = []
    for t in tenants:
        name = t.get("name", "")
        unit = t.get("unit", "")
        phone = t.get("phone", "")
        if q and q not in name.lower() and q not in unit.lower() and q not in phone.lower():
            continue
        txns = []
        for rec in t.get("payment_history", []):
            txns.append({**rec, "kind": "payment"})
        for rec in t.get("deposit_history", []):
            txns.append({**rec, "kind": "deposit"})
        txns.sort(key=lambda r: r.get("date", ""), reverse=True)
        out_txns = []
        for rec in txns:
            frm, to = _compute_txn_period(rec)
            out_txns.append({
                "date": rec.get("date", "—"),
                "kind": rec.get("kind"),
                "cancelled": bool(rec.get("_cancelled", False)),
                "amount": float(rec.get("amount", 0) or 0),
                "from": frm, "to": to,
            })
        out.append({
            "index": t.get("_idx"), "name": name, "unit": unit,
            "entry_date": t.get("entry_date", "") or "—",
            "transactions": out_txns,
        })
    out.sort(key=lambda r: r["name"].lower())
    return jsonify({"tenants": out})


@app.route("/api/units/vacant")
def vacant_units():
    data = load_state()
    occupied = {t.get("unit") for t in data["tenants"] if t.get("unit")}
    exclude = request.args.get("exclude")
    if exclude:
        occupied.discard(exclude)
    out = []
    for name, info in data["units"].items():
        if name in occupied:
            continue
        rent = parse_amount(info.get("rent", 0) if isinstance(info, dict) else info)
        out.append({"name": name, "rent": rent})
    out.sort(key=lambda u: u["name"])
    return jsonify({"units": out})


@app.route("/api/tenants", methods=["POST"])
def add_tenant():
    body = request.get_json(force=True) or {}
    name = (body.get("name") or "").strip().upper()
    phone = (body.get("phone") or "").strip()
    unit = (body.get("unit") or "").strip()
    rent_str = str(body.get("rent") or "").strip()
    entry_str = (body.get("entry_date") or "").strip()
    notes = (body.get("notes") or "").strip()
    replace = bool(body.get("replace"))

    if not name:
        return jsonify({"ok": False, "error": "Tenant name is required."}), 400
    if not unit:
        return jsonify({"ok": False, "error": "Please select a unit."}), 400
    if not phone:
        return jsonify({"ok": False, "error": "Phone number is required."}), 400
    if rent_str and not re.search(r"\d", rent_str):
        return jsonify({"ok": False, "error": "Enter a valid rent amount."}), 400
    try:
        datetime.strptime(entry_str, "%Y-%m-%d")
    except ValueError:
        return jsonify({"ok": False, "error": "Move-in date must be in YYYY-MM-DD format."}), 400

    rent = parse_amount(rent_str) if rent_str else 0.0
    record = {
        "name": name, "phone": phone,
        "email": (body.get("email") or "").strip(),
        "occupation": (body.get("occupation") or "").strip(),
        "emergency_contact": (body.get("emergency_contact") or "").strip(),
        "emergency_phone": (body.get("emergency_phone") or "").strip(),
        "unit": unit, "rent": rent,
        "entry_date": entry_str, "due_date": "",
        "status": "Pending", "pay_date": "", "notes": notes,
        "payment_history": [], "deposit_history": [], "arrears_history": [],
        "locked_periods": [], "deposit_cycle_start": 0, "rent_increase_due": 0.0,
    }

    data = load_state()
    for i, t in enumerate(data["tenants"]):
        if t.get("unit") == unit:
            if not replace:
                return jsonify({"ok": False, "error": "unit_taken",
                                 "message": f"Unit {unit} is already assigned to "
                                            f"{t.get('name', 'Unnamed Tenant')}."}), 409
            data["tenants"][i] = record
            save_state(data)
            return jsonify({"ok": True, "index": i})

    data["tenants"].append(record)
    save_state(data)
    return jsonify({"ok": True, "index": len(data["tenants"]) - 1})


@app.route("/api/tenants/<int:idx>", methods=["PUT"])
def edit_tenant(idx):
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    t = tenants[idx]
    body = request.get_json(force=True) or {}

    editable = ["name", "phone", "email", "occupation", "emergency_contact",
                "emergency_phone", "entry_date", "due_date", "notes"]
    for key in editable:
        if key in body:
            val = str(body[key]).strip()
            t[key] = val.upper() if key == "name" else val
    if "rent" in body:
        rent_str = str(body["rent"]).strip()
        if rent_str and not re.search(r"\d", rent_str):
            return jsonify({"ok": False, "error": "Enter a valid rent amount."}), 400
        t["rent"] = parse_amount(rent_str) if rent_str else 0.0

    save_state(data)
    return jsonify({"ok": True})


@app.route("/api/tenants/<int:idx>", methods=["DELETE"])
def delete_tenant(idx):
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    stamp_tenant_deleted(data, tenants[idx])
    del tenants[idx]
    save_state(data)
    return jsonify({"ok": True})


# ── payments / deposits ───────────────────────────────────────────────────
@app.route("/api/tenants/<int:idx>/payment", methods=["POST"])
def do_record_payment(idx):
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    t = tenants[idx]
    body = request.get_json(force=True) or {}
    try:
        months = int(body.get("months", 1))
        if months < 1:
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Select at least one month."}), 400

    result = record_payment(t, months)
    save_state(data)
    return jsonify({"ok": True, "result": result})


@app.route("/api/tenants/<int:idx>/deposit", methods=["POST"])
def do_record_deposit(idx):
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    t = tenants[idx]
    body = request.get_json(force=True) or {}
    try:
        months = int(body.get("months", 1))
        if months < 1:
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Select at least one month."}), 400
    try:
        instalment = float(str(body.get("amount", "")).strip().replace(",", ""))
        if instalment <= 0:
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Enter a valid deposit amount greater than zero."}), 400

    result = record_deposit(t, months, instalment)
    save_state(data)
    return jsonify({"ok": True, "result": result})


@app.route("/api/tenants/<int:idx>/cancel", methods=["POST"])
def do_cancel(idx):
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    t = tenants[idx]
    body = request.get_json(force=True) or {}
    h_key = body.get("history_key")   # "payment_history" or "deposit_history"
    rec_idx = body.get("record_index")
    if h_key not in ("payment_history", "deposit_history") or rec_idx is None:
        return jsonify({"ok": False, "error": "Invalid record reference."}), 400

    result = cancel_transaction(t, h_key, int(rec_idx))
    if result is None:
        return jsonify({"ok": False, "error": "Record not found."}), 404
    if result == "already_cancelled":
        return jsonify({"ok": False, "error": "That transaction has already been reversed."}), 400
    save_state(data)
    return jsonify({"ok": True, "result": result})


@app.route("/api/tenants/<int:idx>/arrears", methods=["POST"])
def do_clear_arrears(idx):
    data = load_state()
    tenants = data["tenants"]
    if idx < 0 or idx >= len(tenants):
        return jsonify({"error": "not found"}), 404
    t = tenants[idx]
    body = request.get_json(force=True) or {}
    method = body.get("method", "Full")
    due = rent_increase_due(t)
    if due <= 0:
        return jsonify({"ok": False, "error": "No arrears outstanding."}), 400
    try:
        amt = float(str(body.get("amount", "")).strip())
        if amt <= 0:
            raise ValueError
    except (ValueError, TypeError):
        return jsonify({"ok": False, "error": "Enter a valid amount."}), 400
    if amt > due + 0.01:
        return jsonify({"ok": False, "error": f"Amount can't exceed the outstanding UGX {int(due):,}."}), 400
    if method == "Full" and amt < due - 0.01:
        return jsonify({"ok": False, "error": "Full clearance must cover the entire outstanding amount."}), 400

    record_arrears_payment(t, amt, method)
    save_state(data)
    return jsonify({"ok": True})


# ── units ────────────────────────────────────────────────────────────────
@app.route("/api/units")
def list_units():
    data = load_state()
    occupants = {t.get("unit"): t.get("name") for t in data["tenants"]}
    out = []
    for name, info in data["units"].items():
        rent = parse_amount(info.get("rent", 0) if isinstance(info, dict) else info)
        location = info.get("location", "") if isinstance(info, dict) else ""
        pending = info.get("pending_rent_increase") if isinstance(info, dict) else None
        out.append({
            "name": name, "rent": rent, "location": location,
            "occupant": occupants.get(name),
            "pending_rent_increase": pending,
        })
    out.sort(key=lambda u: u["name"])
    return jsonify({"units": out})


@app.route("/api/units", methods=["POST"])
def add_unit():
    body = request.get_json(force=True) or {}
    name = (body.get("name") or "").strip()
    rent_str = str(body.get("rent") or "").strip()
    location = (body.get("location") or "").strip()
    if not name:
        return jsonify({"ok": False, "error": "Enter a unit ID."}), 400
    data = load_state()
    if name in data["units"]:
        return jsonify({"ok": False, "error": f"Unit '{name}' already exists."}), 409
    if rent_str and not re.search(r"\d", rent_str):
        return jsonify({"ok": False, "error": "Enter a valid rent amount."}), 400
    rent = parse_amount(rent_str) if rent_str else 0.0
    data["units"][name] = {"rent": rent, "location": location}
    save_state(data)
    return jsonify({"ok": True})


@app.route("/api/units/<name>", methods=["PUT"])
def edit_unit(name):
    body = request.get_json(force=True) or {}
    data = load_state()
    if name not in data["units"]:
        return jsonify({"ok": False, "error": "Unit not found."}), 404
    rent_str = str(body.get("rent") or "").strip()
    location = (body.get("location") or "").strip()
    if rent_str and not re.search(r"\d", rent_str):
        return jsonify({"ok": False, "error": "Enter a valid rent amount."}), 400
    rent = parse_amount(rent_str) if rent_str else 0.0
    existing = data["units"].get(name, {})
    pending = existing.get("pending_rent_increase") if isinstance(existing, dict) else None
    updated = {"rent": rent, "location": location}
    if pending:
        updated["pending_rent_increase"] = pending
    data["units"][name] = updated
    save_state(data)
    return jsonify({"ok": True})


@app.route("/api/units/<name>", methods=["DELETE"])
def delete_unit(name):
    data = load_state()
    if name not in data["units"]:
        return jsonify({"ok": False, "error": "Unit not found."}), 404
    stamp_unit_deleted(data, name)
    del data["units"][name]
    save_state(data)
    return jsonify({"ok": True})


@app.route("/api/units/<name>/increase-rent", methods=["POST"])
def increase_rent(name):
    body = request.get_json(force=True) or {}
    data = load_state()
    info = data["units"].setdefault(name, {"rent": 0, "location": ""})
    if not isinstance(info, dict):
        info = {"rent": parse_amount(info), "location": ""}
        data["units"][name] = info
    cur_rent = parse_amount(info.get("rent", 0))

    new_rent_str = str(body.get("new_rent") or "").strip()
    if not re.search(r"\d", new_rent_str):
        return jsonify({"ok": False, "error": "Enter a valid rent amount."}), 400
    new_rent = parse_amount(new_rent_str)
    if new_rent <= cur_rent:
        return jsonify({"ok": False,
                         "error": f"New rent must be greater than the current rent of UGX {int(cur_rent):,}."}), 400

    eff_str = (body.get("effective_month") or "").strip()
    try:
        datetime.strptime(eff_str, "%Y-%m-%d")
    except ValueError:
        return jsonify({"ok": False, "error": "Invalid effective month."}), 400

    info["pending_rent_increase"] = {"new_rent": new_rent, "effective_month": eff_str}
    save_state(data)   # load_state (called next time) applies it immediately if due
    data = load_state()  # re-run migrations now so an immediate increase takes effect right away
    return jsonify({"ok": True})


# ── alerts ───────────────────────────────────────────────────────────────
@app.route("/api/alerts")
def alerts():
    data = load_state()
    tenants = _with_index(data["tenants"])
    today = date.today()
    items = get_watchlist_tenants(tenants, max_days=36500)
    out = [_tenant_summary(t, today) | {"days_left": d} for t, d in items]
    return jsonify({"alerts": out})


# ── exports ──────────────────────────────────────────────────────────────
@app.route("/api/monthly-report")
def monthly_report():
    try:
        year = int(request.args.get("year", date.today().year))
        month = int(request.args.get("month", date.today().month))
        if not (1 <= month <= 12):
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "Invalid year/month."}), 400
    data = load_state()
    report = build_monthly_report(data, year, month)
    return jsonify(report)


@app.route("/api/export/monthly-excel")
def do_export_monthly_excel():
    try:
        year = int(request.args.get("year", date.today().year))
        month = int(request.args.get("month", date.today().month))
        if not (1 <= month <= 12):
            raise ValueError
    except (TypeError, ValueError):
        return jsonify({"error": "Invalid year/month."}), 400
    data = load_state()
    report = build_monthly_report(data, year, month)
    path = export_monthly_excel(report)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/api/export/excel")
def do_export_excel():
    data = load_state()
    path = export_excel(data)
    return send_file(path, as_attachment=True, download_name="tenant_records.xlsx")


@app.route("/api/export/pdf")
def do_export_pdf():
    data = load_state()
    path = export_pdf(data)
    return send_file(path, as_attachment=True, download_name="tenant_data.pdf")


# ── settings / data management ────────────────────────────────────────────
@app.route("/api/shutdown", methods=["POST"])
def shutdown_server():
    """Gracefully stops this process. Runs independently of the desktop
    GUI (see module docstring), so the desktop app calls this over HTTP
    to stop it — via 'Stop Phone Access' — rather than needing to hold a
    direct handle to this process, which it may not have if the web
    companion was started in an earlier desktop-app session and is still
    running after that session's window was closed."""
    def _die():
        import time as _t
        _t.sleep(0.3)  # give the HTTP response below a moment to flush
        os._exit(0)
    threading.Thread(target=_die, daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/settings/reset", methods=["POST"])
def reset_data():
    """Wipes all tenant/unit data. Requires the admin PIN that was
    originally set on the PC (settings.pin_hash -- the same one used to
    lock/unlock this app) to be re-entered, so a device that merely has
    an open, unlocked session can't nuke everything on its own; a backup
    of the current data is always taken first regardless."""
    data = load_state()
    pin_hash = data.get("settings", {}).get("pin_hash", "")
    if not pin_hash:
        return jsonify({"ok": False,
                         "error": "No admin PIN is set on the PC yet. Set one from "
                                  "the desktop app's Settings before resetting data "
                                  "from here."}), 400
    body = request.get_json(force=True) or {}
    entered_pin = (body.get("admin_pin") or "").strip()
    if not entered_pin or hash_secret(entered_pin) != pin_hash:
        return jsonify({"ok": False, "error": "Incorrect admin PIN."}), 400
    backup_current_data_file()
    save_state({"units": {}, "tenants": [], "settings": {}})
    session["unlocked"] = True
    return jsonify({"ok": True})

# =========================================================================
#  SECTION 4 -- Mobile frontend (single-page app, embedded as one string)
# =========================================================================
INDEX_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1, viewport-fit=cover">
<title>Tenant Management</title>
<!-- PWA: lets this page be saved/installed to a phone home screen or
     desktop browser (Chrome/Edge install icon, Safari "Add to Home
     Screen") so it launches full-screen with its own icon, for easy
     repeat access without re-typing the LAN address each time. -->
<link rel="manifest" href="manifest.json">
<meta name="theme-color" content="#0E4F4F">
<link rel="icon" href="icon-192.png">
<link rel="apple-touch-icon" href="icon-256.png">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Tenant Management">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
  :root{
    --teal-deep:#0E4F4F;
    --teal:#1FAD9F;
    --teal-soft:#EAF7F4;
    --teal-soft2:#F0FAF8;
    --ink:#0E2222;
    --muted:#6E8482;
    --line:#DCEEEA;
    --card:#FFFFFF;
    --bg:#F5FAF9;
    --danger:#EF4565;
    --warn:#FFB020;
    --good:#3DBE7C;
    --radius:16px;
  }
  *{box-sizing:border-box;}
  html,body{margin:0;padding:0;}
  body{
    background:var(--bg);
    color:var(--ink);
    font-family:'Inter',system-ui,sans-serif;
    -webkit-font-smoothing:antialiased;
    padding-bottom:84px;
    min-height:100vh;
  }
  h1,h2,h3,.disp{font-family:'Space Grotesk',system-ui,sans-serif;}
  .app{max-width:520px;margin:0 auto;position:relative;}
  header.top{
    position:sticky;top:0;z-index:20;
    background:linear-gradient(180deg,var(--teal-deep),#0B3E3E);
    color:#fff;padding:18px 18px 16px;
    display:flex;align-items:center;justify-content:space-between;
  }
  header.top .brand{display:flex;align-items:center;gap:10px;}
  header.top .brand .mark{
    width:34px;height:34px;border-radius:10px;background:var(--teal);
    display:flex;align-items:center;justify-content:center;font-weight:700;
  }
  header.top h1{font-size:16px;margin:0;font-weight:600;letter-spacing:.2px;}
  header.top .sub{font-size:11px;opacity:.75;margin-top:1px;}
  .icon-btn{
    width:38px;height:38px;border-radius:12px;background:rgba(255,255,255,.12);
    border:none;color:#fff;display:flex;align-items:center;justify-content:center;
    font-size:16px;cursor:pointer;
  }
  .icon-btn:disabled{opacity:.5;cursor:default;}
  main{padding:16px;}
  .card{
    background:var(--card);border-radius:var(--radius);
    box-shadow:0 1px 2px rgba(14,79,79,.06),0 8px 24px -16px rgba(14,79,79,.25);
    padding:16px;margin-bottom:14px;border:1px solid var(--line);
  }
  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
  .stat{padding:14px;border-radius:14px;background:var(--teal-soft2);}
  .stat .num{font-family:'Space Grotesk';font-size:22px;font-weight:700;color:var(--teal-deep);}
  .stat .lbl{font-size:11.5px;color:var(--muted);margin-top:2px;}

  /* ── Dashboard stat cards — mirrors the desktop app's Dashboard tab ── */
  .dgrid{display:grid;grid-template-columns:1fr;gap:12px;margin-top:2px;}
  .dcard{
    border-radius:18px;overflow:hidden;display:flex;flex-direction:column;
    box-shadow:0 1px 2px rgba(14,79,79,.06),0 8px 24px -16px rgba(14,79,79,.25);
  }
  .dcard-top{padding:16px 14px 14px;color:#fff;position:relative;}
  .dcard-icon{position:absolute;top:10px;right:12px;font-size:26px;opacity:.9;}
  .dcard-title{font-size:11px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;opacity:.92;padding-right:30px;}
  .dcard-value{font-family:'Space Grotesk';font-size:24px;font-weight:700;margin-top:6px;line-height:1.1;word-break:break-word;}
  .dcard-footer{background:#fff;padding:10px 12px 0;}
  .dcard-subrow{display:flex;gap:8px;padding-bottom:8px;}
  .dcard-sub{flex:1;min-width:0;}
  .dcard-sub .v{font-family:'Space Grotesk';font-weight:700;font-size:13px;word-break:break-word;}
  .dcard-sub .l{font-size:10px;color:var(--muted);margin-top:1px;}
  .dcard-action{
    display:block;border-top:1px solid var(--line);text-align:right;
    padding:10px 4px;font-size:13px;font-weight:700;background:none;border-left:none;border-right:none;border-bottom:none;
    width:100%;cursor:pointer;
  }
  .dcard-action:active{background:var(--teal-soft2);}
  .btn{
    border:none;border-radius:12px;font-weight:600;font-size:14px;
    padding:12px 16px;cursor:pointer;display:inline-flex;align-items:center;gap:8px;justify-content:center;
  }
  .btn-primary{background:var(--teal);color:#fff;}
  .btn-primary:active{background:var(--teal-deep);}
  .btn-ghost{background:var(--teal-soft);color:var(--teal-deep);}
  .btn-danger{background:#FCE8EC;color:var(--danger);}
  .btn-full{width:100%;}
  .btn:disabled{opacity:.5;}
  input,select,textarea{
    width:100%;border:1.5px solid var(--line);border-radius:12px;
    padding:11px 12px;font-size:15px;font-family:inherit;background:#fff;color:var(--ink);
  }
  input:focus,select:focus,textarea:focus{outline:none;border-color:var(--teal);}
  label.field{display:block;font-size:12.5px;color:var(--muted);font-weight:600;margin:10px 0 5px;}
  .row{display:flex;gap:10px;}
  .row > *{flex:1;}
  .chip{
    display:inline-flex;align-items:center;gap:5px;font-size:12px;font-weight:600;
    padding:5px 10px;border-radius:999px;
  }
  .chip-paid{background:#E4F8EE;color:var(--good);}
  .chip-underpaid{background:#FFF4E0;color:#B5790A;}
  .chip-pending{background:#FCE8EC;color:var(--danger);}
  .tenant-row{
    display:flex;align-items:center;gap:12px;padding:12px 4px;border-bottom:1px solid var(--line);cursor:pointer;
  }
  .tenant-row:last-child{border-bottom:none;}
  .avatar{
    width:42px;height:42px;border-radius:12px;background:var(--teal-soft);color:var(--teal-deep);
    display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;flex-shrink:0;
  }
  .tenant-row .meta{flex:1;min-width:0;}
  .tenant-row .name{font-weight:600;font-size:14.5px;}
  .tenant-row .sub{font-size:12px;color:var(--muted);margin-top:1px;}
  .searchbar{
    display:flex;align-items:center;gap:8px;background:#fff;border:1.5px solid var(--line);
    border-radius:12px;padding:9px 12px;margin-bottom:12px;
  }
  .menu-row{
    display:flex;align-items:center;gap:14px;padding:16px 18px;cursor:pointer;
    border-bottom:1px solid var(--line);
  }
  .menu-row:last-child{border-bottom:none;}
  .menu-row:active{background:var(--teal-soft2);}
  .menu-icon{font-size:22px;flex-shrink:0;}
  .menu-text{flex:1;min-width:0;}
  .menu-title{font-weight:700;font-size:15px;}
  .menu-sub{font-size:12px;color:var(--muted);margin-top:2px;}
  .menu-chevron{font-size:20px;color:var(--muted);}
  .hist-empty{
    background:var(--teal-soft2);border-radius:12px;padding:12px 14px;
    font-size:12.5px;color:var(--muted);
  }
  .hist-table{border:1px solid var(--line);border-radius:12px;overflow:hidden;}
  .hist-hdr{
    display:grid;grid-template-columns:1.1fr 1fr .9fr 1.4fr;gap:4px;
    background:var(--teal-soft2);padding:8px 10px;font-size:10px;font-weight:700;
    letter-spacing:.3px;text-transform:uppercase;color:var(--muted);
  }
  .hist-row{
    display:grid;grid-template-columns:1.1fr 1fr .9fr 1.4fr;gap:4px;
    padding:9px 10px;font-size:12px;border-top:1px solid var(--line);align-items:center;
  }
  .hist-row.hist-cancelled{background:#FFF0F0;color:var(--danger);}
  .hist-period{color:var(--muted);font-size:11px;}
  .hist-row.hist-cancelled .hist-period{color:var(--danger);opacity:.8;}
  .month-picker-control{
    width:100%;border:1.5px solid var(--line);border-radius:12px;padding:11px 12px;
    font-size:14px;background:#fff;color:var(--ink);display:flex;align-items:center;
    justify-content:space-between;cursor:pointer;
  }
  .month-picker-control .mp-caret{color:var(--muted);font-size:12px;}
  .select-native-wrap{position:relative;flex:1;}
  .select-native-wrap select{
    width:100%;appearance:none;-webkit-appearance:none;-moz-appearance:none;
    border:1.5px solid var(--line);border-radius:12px;padding:11px 30px 11px 12px;
    font-size:14px;background:#fff;color:var(--ink);
  }
  .select-native-wrap::after{
    content:'▾';position:absolute;right:12px;top:50%;transform:translateY(-50%);
    color:var(--muted);font-size:12px;pointer-events:none;
  }
  .month-picker-panel{
    border:1.5px solid var(--line);border-radius:12px;margin-top:6px;overflow:hidden;background:#fff;
  }
  .month-picker-list{max-height:230px;overflow-y:auto;}
  .month-row{
    display:flex;align-items:center;gap:10px;padding:10px 12px;font-size:13.5px;
    border-bottom:1px solid var(--line);cursor:pointer;
  }
  .month-row:last-child{border-bottom:none;}
  .month-row input[type=checkbox]{width:17px;height:17px;flex-shrink:0;pointer-events:none;}
  .month-row .month-label{flex:1;}
  .month-row.cleared{background:var(--teal-soft2);color:var(--muted);cursor:default;}
  .month-row.cleared .month-tag{
    font-size:10px;font-weight:700;color:var(--teal-deep);background:var(--teal-soft);
    padding:2px 7px;border-radius:999px;text-transform:uppercase;letter-spacing:.3px;
  }
  .month-row.open:active{background:var(--teal-soft2);}
  .month-row.open.ticked{background:var(--teal-soft2);}
  .month-picker-actions{
    display:flex;gap:8px;padding:10px 12px;border-top:1px solid var(--line);background:var(--teal-soft2);
  }
  .month-picker-actions .btn{padding:9px 14px;font-size:13px;}
  .searchbar input{border:none;padding:0;font-size:14px;}
  .filters{display:flex;gap:8px;overflow-x:auto;padding-bottom:2px;margin-bottom:14px;}
  .filter-pill{
    flex-shrink:0;padding:7px 14px;border-radius:999px;background:#fff;border:1.5px solid var(--line);
    font-size:12.5px;font-weight:600;color:var(--muted);cursor:pointer;
  }
  .filter-pill.active{background:var(--teal-deep);border-color:var(--teal-deep);color:#fff;}
  nav.tabbar{
    position:fixed;bottom:0;left:0;right:0;background:#fff;border-top:1px solid var(--line);
    display:flex;z-index:30;padding-bottom:env(safe-area-inset-bottom);
  }
  nav.tabbar .tab{
    flex:1;padding:10px 4px 8px;display:flex;flex-direction:column;align-items:center;gap:3px;
    color:var(--muted);font-size:10.5px;font-weight:600;cursor:pointer;
  }
  nav.tabbar .tab.active{color:var(--teal-deep);}
  nav.tabbar .tab .ic{font-size:19px;}
  .empty{text-align:center;padding:36px 12px;color:var(--muted);}
  .empty .big{font-size:32px;margin-bottom:8px;}
  .section-title{font-size:13px;font-weight:700;color:var(--teal-deep);text-transform:uppercase;letter-spacing:.4px;margin:4px 0 10px;}
  .modal-backdrop{
    position:fixed;inset:0;background:rgba(10,30,30,.5);z-index:100;
    display:flex;align-items:flex-end;justify-content:center;
  }
  .modal{
    background:#fff;border-radius:20px 20px 0 0;width:100%;max-width:520px;max-height:88vh;overflow-y:auto;
    padding:20px 18px calc(20px + env(safe-area-inset-bottom));animation:slideup .18s ease-out;
  }
  @keyframes slideup{from{transform:translateY(24px);opacity:0;}to{transform:translateY(0);opacity:1;}}
  .modal h2{font-size:17px;margin:0 0 4px;}
  .modal .desc{font-size:13px;color:var(--muted);margin-bottom:14px;line-height:1.4;}
  .modal .close-x{position:absolute;top:14px;right:14px;}
  .err{color:var(--danger);font-size:12.5px;margin-top:8px;min-height:1px;}
  .toast{
    position:fixed;bottom:92px;left:50%;transform:translateX(-50%);
    background:var(--teal-deep);color:#fff;padding:11px 18px;border-radius:12px;font-size:13.5px;
    z-index:200;max-width:88%;text-align:center;box-shadow:0 8px 24px rgba(0,0,0,.2);
  }
  .lockscreen{
    position:fixed;inset:0;background:linear-gradient(160deg,var(--teal-deep),#0a3535);
    z-index:1000;display:flex;flex-direction:column;align-items:center;justify-content:center;color:#fff;padding:24px;
  }
  .pin-dots{display:flex;gap:12px;margin:22px 0;}
  .pin-dots .d{width:14px;height:14px;border-radius:50%;border:2px solid rgba(255,255,255,.5);}
  .pin-dots .d.filled{background:#fff;border-color:#fff;}
  .keypad{display:grid;grid-template-columns:repeat(3,64px);gap:14px;margin-top:10px;}
  .keypad button{
    width:64px;height:64px;border-radius:50%;background:rgba(255,255,255,.1);border:none;color:#fff;
    font-size:20px;font-weight:600;cursor:pointer;
  }
  .keypad button:active{background:rgba(255,255,255,.25);}
  .disc-screen{
    position:fixed;inset:0;background:linear-gradient(160deg,var(--teal-deep),#0a3535);
    z-index:1100;display:flex;flex-direction:column;align-items:center;justify-content:center;
    color:#fff;padding:32px;text-align:center;
  }
  .disc-screen .icon{font-size:52px;margin-bottom:14px;}
  .disc-screen h1{margin:0 0 10px;font-size:21px;}
  .disc-screen p{margin:0;font-size:13.5px;opacity:.85;max-width:320px;line-height:1.5;}
  .disc-screen .spin{
    margin-top:22px;width:20px;height:20px;border-radius:50%;
    border:2.5px solid rgba(255,255,255,.3);border-top-color:#fff;
    animation:disc-spin 0.9s linear infinite;
  }
  @keyframes disc-spin{to{transform:rotate(360deg);}}
  .txn-item{display:flex;justify-content:space-between;align-items:center;padding:10px 0;border-bottom:1px solid var(--line);}
  .txn-item:last-child{border-bottom:none;}
  .txn-item .amt{font-weight:700;font-size:14px;}
  .txn-item .amt.cancelled{text-decoration:line-through;color:var(--muted);font-weight:600;}
  .badge-dot{
    position:absolute;top:-4px;right:-4px;background:var(--danger);color:#fff;font-size:10px;font-weight:700;
    min-width:16px;height:16px;border-radius:8px;display:flex;align-items:center;justify-content:center;padding:0 3px;
  }
  a.linklike{color:var(--teal-deep);font-weight:600;text-decoration:none;font-size:13px;}
  .progress-bar{height:8px;border-radius:6px;background:var(--teal-soft);overflow:hidden;margin-top:6px;}
  .progress-bar > div{height:100%;background:var(--teal);}
  hr.sep{border:none;border-top:1px solid var(--line);margin:14px 0;}
  .sync-badge{
    font-size:11px;font-weight:700;padding:5px 10px;border-radius:999px;
    background:rgba(255,255,255,.14);color:#fff;white-space:nowrap;margin-right:8px;
  }
  .sync-badge.offline{background:#FFB020;color:#3A2900;}
  .sync-badge.syncing{background:#2AABEE;color:#fff;}
  .sync-badge.online{background:#1E9E6B;color:#fff;}
  .sync-badge.disconnected{background:#D93B3B;color:#fff;}
  .ptr-indicator{
    text-align:center;font-size:12px;color:var(--muted);font-weight:700;
    height:0;overflow:hidden;transition:height .15s ease;
  }
  .ptr-indicator.visible{height:34px;line-height:34px;}
  .ptr-indicator.ready{color:var(--teal-deep);}
  .ptr-indicator.spinning{animation:ptr-pulse 1s ease-in-out infinite;}
  @keyframes ptr-pulse{0%,100%{opacity:.5;}50%{opacity:1;}}
  .pending-note{
    background:#FFF4E0;color:#8A5A00;border:1px solid #F3D9A0;border-radius:12px;
    padding:10px 12px;font-size:12.5px;font-weight:600;margin-bottom:12px;
  }
</style>
</head>
<body>

<div id="lockscreen" class="lockscreen" style="display:none;">
  <div style="font-size:15px;opacity:.85;">Tenant Management</div>
  <h1 style="margin:4px 0 0;font-size:20px;" id="lockTitle">Enter PIN</h1>
  <div class="pin-dots" id="pinDots"></div>
  <div class="err" id="lockErr" style="color:#FFB4C0;min-height:16px;"></div>
  <div class="keypad" id="keypad"></div>
</div>

<div id="discscreen" class="disc-screen" style="display:none;">
  <div class="icon" id="discIcon">🔌</div>
  <h1 id="discTitle">Disconnected</h1>
  <p id="discMsg">Sharing was turned off on the PC, so no tenant data is available on this phone right now. This closes automatically once it's reconnected.</p>
  <div class="spin"></div>
</div>

<div id="devlimitscreen" class="disc-screen" style="display:none;">
  <div class="icon">📵</div>
  <h1>Device Limit Reached</h1>
  <p id="devlimitMsg">Up to 3 phones can be connected at once, and that limit is currently full. This connects automatically as soon as one of the other devices disconnects — no need to scan again.</p>
  <div class="spin"></div>
</div>

<div class="app" id="app" style="display:none;">
  <header class="top">
    <div class="brand">
      <div class="mark">🏠</div>
      <div>
        <h1>Tenant Management</h1>
        <div class="sub" id="headerSub">—</div>
      </div>
    </div>
    <div id="syncBadge" class="sync-badge" style="display:none;"></div>
    <button class="icon-btn" onclick="lockNow()" title="Lock">🔒</button>
  </header>
  <div id="ptrIndicator" class="ptr-indicator">↓ Pull to refresh</div>
  <main id="main"></main>
</div>

<nav class="tabbar" id="tabbar" style="display:none;">
  <div class="tab" data-tab="dashboard"><span class="ic">📊</span>Home</div>
  <div class="tab" data-tab="tenants"><span class="ic">👥</span>Tenants</div>
  <div class="tab" data-tab="units"><span class="ic">🏢</span>Units</div>
  <div class="tab" data-tab="alerts" style="position:relative;"><span class="ic">🔔</span>Alerts<span class="badge-dot" id="alertDot" style="display:none;"></span></div>
  <div class="tab" data-tab="more"><span class="ic">⚙️</span>More</div>
</nav>

<div id="modalRoot"></div>
<div id="toastRoot"></div>

<script>
const $ = (sel, el=document) => el.querySelector(sel);
const $$ = (sel, el=document) => [...el.querySelectorAll(sel)];
const fmt = n => 'UGX ' + Math.round(n||0).toLocaleString();
const todayStr = () => new Date().toISOString().slice(0,10);

function parseISO(s) {
  if (!s) return null;
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (!m) return null;
  return new Date(Date.UTC(+m[1], +m[2]-1, +m[3]));
}
function addMonthsJS(d, n) {
  const day = d.getUTCDate();
  const nd = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth()+n, 1));
  const lastDay = new Date(Date.UTC(nd.getUTCFullYear(), nd.getUTCMonth()+1, 0)).getUTCDate();
  nd.setUTCDate(Math.min(day, lastDay));
  return nd;
}
function diffDays(a, b) { return Math.round((a - b) / 86400000); }

// Mirrors the desktop app's "Lease Summary" card: a progress bar from
// last-due-date to next-due-date while paid, or a plain "Days Overdue"
// figure once pending — same math as ModernRentalApp's tenant detail.
function leaseProgressBlock(t) {
  const today = parseISO(todayStr());
  if (t.level === 'pending') {
    const refD = parseISO(t.due_date || t.entry_date);
    const overdueTxt = refD ? `${diffDays(today, refD)} day(s)` : '—';
    return `<div class="card">
      <div class="section-title" style="margin-top:0;">Days Overdue</div>
      <div style="font-size:24px;font-weight:700;color:var(--danger);">${overdueTxt}</div>
    </div>`;
  }
  const dueD = parseISO(t.due_date);
  if (!dueD) return '';
  const fromD    = addMonthsJS(dueD, -1);
  const remDays  = diffDays(dueD, today);
  const totalDays = diffDays(dueD, fromD) || 1;
  const elapsed   = diffDays(today, fromD);
  const pct       = Math.max(0, Math.min(1, elapsed / totalDays));
  const pctPx     = Math.round(pct * 100);
  const daysNum   = Math.abs(remDays);
  const circSize  = daysNum < 100 ? '22px' : '15px';
  return `<div class="card">
    <div class="section-title" style="margin-top:0;">Lease Progress</div>
    <div style="display:flex;align-items:center;gap:16px;">
      <div style="flex:1;min-width:0;">
        <div style="font-size:13px;font-weight:700;margin-bottom:6px;">~${pctPx}%</div>
        <div class="progress-bar"><div style="width:${pctPx}%"></div></div>
        <div style="display:flex;justify-content:space-between;font-size:11px;color:var(--muted);margin-top:6px;">
          <span>↑ Start: ${fromD.toISOString().slice(0,10)}</span>
          <span>↑ End: ${dueD.toISOString().slice(0,10)}</span>
        </div>
      </div>
      <div style="text-align:center;flex-shrink:0;">
        <div style="width:64px;height:64px;border-radius:50%;border:3px solid var(--line);
                    display:flex;align-items:center;justify-content:center;font-weight:700;
                    font-size:${circSize};color:${remDays<0?'var(--danger)':'var(--ink)'};">${daysNum}</div>
        <div style="font-size:10px;font-weight:700;color:var(--muted);margin-top:4px;">${remDays<0?'OVERDUE':'DAYS LEFT'}</div>
      </div>
    </div>
  </div>`;
}

let state = { tab:'dashboard', tenants:[], selectedIdx:null, filter:'all', q:'' };

// ── Offline queue + cache ───────────────────────────────────────────
// While this phone/browser can't reach the PC (PC is off, or off Wi-Fi),
// reads are served from the last-known-good copy of each GET response,
// and any Add/Edit/Payment/etc. is saved to a local queue instead of
// failing outright. As soon as the PC is reachable again, the queue is
// replayed against the real backend automatically, in the order it was
// made, and the screen refreshes with the authoritative server data.
const CACHE_KEY = 'rm_offline_cache_v1';
const QUEUE_KEY = 'rm_offline_queue_v1';
const DEVICE_ID_KEY = 'rm_device_id_v1';
function getDeviceId() {
  let id = localStorage.getItem(DEVICE_ID_KEY);
  if (!id) {
    id = (crypto.randomUUID ? crypto.randomUUID() : (Date.now() + '-' + Math.random().toString(16).slice(2)));
    try { localStorage.setItem(DEVICE_ID_KEY, id); } catch(e) {}
  }
  return id;
}
let DEVICE_ID = getDeviceId();
function adoptCanonicalDeviceId(canonical) {
  // The server recognizes phones primarily by this localStorage-held id,
  // but that storage isn't always the same across contexts on the SAME
  // physical phone: iOS treats a web app "Added to Home Screen" as a
  // separate storage silo from Safari itself, and some in-app/QR-scanner
  // browsers do too. When the server's fingerprint match (see
  // DEVICE_FINGERPRINT below) says this context is really an
  // already-known device, it hands back that device's real id -- adopt
  // it here so this context IS that device from now on, not a new one.
  if (!canonical || canonical === DEVICE_ID) return;
  DEVICE_ID = canonical;
  try { localStorage.setItem(DEVICE_ID_KEY, canonical); } catch(e) {}
}
// Safari deliberately omits the specific iPhone/iPad model from its
// User-Agent string (unlike Android, which usually names the exact
// model), so there's no reliable way to ask the browser "which iPhone is
// this". Physical screen resolution is the best available hint -- not
// perfect (several generations share the same screen size) but enough
// for the server to make a reasonable guess instead of just "iPhone".
const DEVICE_SCREEN_HINT = `${screen.width}x${screen.height}@${window.devicePixelRatio || 1}`;
// A small, stable hash of signals that stay the same for this physical
// device/browser combination regardless of which storage silo a given
// page load happens to be in (Safari tab vs. home-screen install vs. an
// in-app browser) -- used server-side to recognize "this is the same
// phone as an already-known device" even when localStorage itself
// isn't shared between those contexts. Not meant to be unique across
// ALL devices (two identical phone models could coincide) -- just
// stable for THIS one across contexts, which is what actually matters
// here: the fallback only ever kicks in for an otherwise-unrecognized
// device_id, so a coincidental match just means one fewer duplicate
// entry on the household's device list, never a security boundary.
function _fnv1aHash(str) {
  let h = 0x811c9dc5;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    h = Math.imul(h, 0x01000193);
  }
  return (h >>> 0).toString(16);
}
const DEVICE_FINGERPRINT = _fnv1aHash([
  navigator.userAgent,
  screen.width + 'x' + screen.height,
  navigator.platform || '',
  navigator.hardwareConcurrency || '',
  navigator.maxTouchPoints || '',
  navigator.language || '',
].join('|'));
let isOnline = true;
let syncing = false;
let tempIdCounter = 0;

function loadCache() { try { return JSON.parse(localStorage.getItem(CACHE_KEY)) || {}; } catch(e) { return {}; } }
function saveCache(c) { try { localStorage.setItem(CACHE_KEY, JSON.stringify(c)); } catch(e) {} }
function cacheGet(path) { return loadCache()[path]; }
function cacheSet(path, data) { const c = loadCache(); c[path] = data; saveCache(c); }
function cacheDeletePrefix(prefix) {
  const c = loadCache();
  Object.keys(c).forEach(k => { if (k.indexOf(prefix) === 0) delete c[k]; });
  saveCache(c);
}
function loadQueue() { try { return JSON.parse(localStorage.getItem(QUEUE_KEY)) || []; } catch(e) { return []; } }
function saveQueue(q) { try { localStorage.setItem(QUEUE_KEY, JSON.stringify(q)); } catch(e) {} updateSyncBadge(); }

function updateSyncBadge() {
  const el = $('#syncBadge');
  if (!el) return;
  const pending = loadQueue().length;
  el.style.display = 'inline-block';
  if (syncing) {
    el.className = 'sync-badge syncing';
    el.textContent = '⟳ Syncing…';
  } else if (explicitDisconnect) {
    // The PC deliberately turned sharing off (Settings → Disconnect) --
    // distinct from a plain network drop, per the three states this
    // badge is meant to represent (Online / Offline / Disconnected).
    el.className = 'sync-badge disconnected';
    el.textContent = '⛔ Disconnected';
  } else if (!isOnline) {
    el.className = 'sync-badge offline';
    el.textContent = pending ? `📴 Offline · ${pending} pending` : '📴 Offline';
  } else {
    el.className = 'sync-badge online';
    el.textContent = pending ? `🟢 Online · ${pending} pending` : '🟢 Online';
  }
}

let booted = false;
let explicitDisconnect = false;
let wasKicked = false;
let wasPending = false;

// Four distinct "not showing you data right now" states:
//
// 1. Explicit disconnect — the desktop app tapped Settings → Disconnect
//    and told this server so (see /api/announce-disconnect) *before*
//    actually shutting it down, so this phone catches it on its next
//    poll while still connected. Blocking screen, no tenant data shown,
//    matching what the PC is telling every phone: sharing is off.
// 2. Kicked — the desktop app's Connect Phone device list disconnected
//    just THIS device (see /api/devices/<id>/kick) while everyone else
//    stays connected. Also blocking, but won't clear itself the way
//    explicit-disconnect does: this device id is permanently blocked
//    server-side, so it needs a fresh scan to come back.
// 3. Pending approval — reaching this page at all (QR scan or a
//    forwarded link — the server genuinely cannot tell them apart) is
//    NOT enough on its own. Every new device id lands here first and
//    stays blocked, showing no tenant data whatsoever, until someone
//    with the PC in front of them explicitly approves it from Settings
//    → Connect Phone. This is what makes sharing the link pointless:
//    the recipient just sees "Waiting for approval" forever unless the
//    PC owner deliberately lets that specific device in.
// 4. Network/PC loss — the server just stopped answering, with no such
//    warning (PC powered off, lost Wi-Fi/internet, etc). Non-blocking:
//    keep showing the last-known tenant data from cache, keep the app
//    fully usable, and queue any changes to sync automatically once
//    reachable again. See the offline queue + cache section above.
function setOnline(v) {
  const wasOffline = !isOnline;
  isOnline = v;
  if (v) {
    explicitDisconnect = false;
    wasKicked = false;
    wasPending = false;
    hideDisconnected();
    if (wasOffline) flushQueue();
  }
  updateSyncBadge();
}

function handleExplicitDisconnect() {
  isOnline = false;
  explicitDisconnect = true;
  wasKicked = false;
  wasPending = false;
  updateSyncBadge();
  $('#discIcon').textContent = '🔌';
  $('#discTitle').textContent = 'Disconnected';
  $('#discMsg').textContent = "Sharing was turned off on the PC, so no tenant data is available on this phone right now. This closes automatically once it's reconnected.";
  showDisconnected();
}

function handleKicked() {
  isOnline = false;
  wasKicked = true;
  explicitDisconnect = false;
  wasPending = false;
  updateSyncBadge();
  $('#discIcon').textContent = '🚫';
  $('#discTitle').textContent = 'Disconnected by Admin';
  $('#discMsg').textContent = "This device was disconnected from the PC's Connect Phone list. It won't reconnect automatically — scan the QR code again on the PC if access should be restored.";
  showDisconnected();
}

function handlePendingApproval() {
  isOnline = false;
  wasPending = true;
  explicitDisconnect = false;
  wasKicked = false;
  updateSyncBadge();
  $('#discIcon').textContent = '🔐';
  $('#discTitle').textContent = 'Waiting for Approval';
  $('#discMsg').textContent = "Opening this link isn't enough on its own — someone needs to approve this device from Settings → Connect Phone on the PC before any tenant data is shown here. This closes automatically once approved.";
  showDisconnected();
}

function showDisconnected() {
  if (explicitDisconnect || wasKicked || wasPending) { $('#discscreen').style.display = 'flex'; }
}
function hideDisconnected() {
  $('#discscreen').style.display = 'none';
}

function showDeviceLimit(maxDevices) {
  $('#lockscreen').style.display = 'none';
  $('#app').style.display = 'none';
  $('#tabbar').style.display = 'none';
  $('#discscreen').style.display = 'none';
  const n = maxDevices || 3;
  $('#devlimitMsg').textContent =
    `Up to ${n} phones can be connected at once, and that limit is currently full. This connects automatically as soon as one of the other devices disconnects — no need to scan again.`;
  $('#devlimitscreen').style.display = 'flex';
}
function hideDeviceLimit() {
  $('#devlimitscreen').style.display = 'none';
}

// Single retry loop behind every blocking state above (kicked, explicit
// disconnect, pending approval, device limit) — re-checks lock-status
// every 5s and either re-shows whichever blocking screen still applies,
// or clears everything and proceeds to the lock screen / app once
// whatever was blocking has genuinely resolved.
let blockedRetryTimer = null;
function enterBlockedState(ls) {
  if (ls.kicked) { handleKicked(); }
  else if (ls.disconnecting) { handleExplicitDisconnect(); }
  else if (ls.pending_approval) { handlePendingApproval(); }
  else if (ls.device_limit_reached) { showDeviceLimit(ls.max_devices); }
  else { return; }
  scheduleBlockedRetry();
}
function scheduleBlockedRetry() {
  if (blockedRetryTimer) return; // already waiting on one
  blockedRetryTimer = setTimeout(async () => {
    blockedRetryTimer = null;
    try {
      const res = await fetchTimeout('/api/lock-status', {headers: {'X-Device-Id': DEVICE_ID}}, 3500);
      const ls = await res.json();
      if (ls && (ls.kicked || ls.disconnecting || ls.pending_approval || ls.device_limit_reached)) {
        enterBlockedState(ls);
        return;
      }
      hideDeviceLimit();
      hideDisconnected();
      cacheSet('/api/lock-status', ls);
      setOnline(true);
      if (ls.pin_set && !ls.unlocked) { showLock(); } else { hideLock(); boot(); }
      updateSyncBadge();
    } catch (e) {
      scheduleBlockedRetry();
    }
  }, 5000);
}

// Fetch with a short timeout — an unreachable LAN IP can otherwise hang
// for a very long time before the browser gives up on its own.
async function fetchTimeout(path, opts={}, ms=4000) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), ms);
  // Every call site already sends X-Device-Id; adding the screen hint
  // and fingerprint here once means every request tags along the
  // model-guessing hint and the cross-context device-recognition signal
  // without having to touch every fetchTimeout(...) call individually.
  const headers = { ...(opts.headers || {}), 'X-Device-Screen': DEVICE_SCREEN_HINT, 'X-Device-Fingerprint': DEVICE_FINGERPRINT };
  if (CLOUD_DIRECT && path.charAt(0) === '/' && path.indexOf('/api/') === 0) {
    headers['X-Session-Id'] = CLOUD_DIRECT.sessionId;
    headers['X-Secret-Key'] = CLOUD_DIRECT.secretKey;
  }
  try {
    const res = await fetch(path, {...opts, headers, signal: ctrl.signal});
    adoptCanonicalDeviceId(res.headers.get('X-Canonical-Device-Id'));
    return res;
  } finally {
    clearTimeout(timer);
  }
}

// Skipped whenever a refresh could disrupt someone actively typing/
// editing, or when there's nothing to refresh anyway. Shared by both the
// slow 90s baseline timer below and the fast change-detection in
// pingServer(), so "don't interrupt what the person is doing" is defined
// in exactly one place.
function refreshIfSafe() {
  if (document.hidden) return;
  const lockEl = $('#lockscreen');
  if (lockEl && lockEl.style.display !== 'none') return;
  const modalEl = $('#modalRoot');
  if (modalEl && modalEl.innerHTML.trim() !== '') return;
  if (state.tab === 'add-tenant') return;
  render();
}

let pingMisses = 0;
// Last data_updated_at we've seen from the server (via /api/lock-status).
// null until the first successful ping, so that first ping only primes
// this value instead of firing a refresh (there's nothing "new" about
// data the app hasn't loaded yet).
let lastKnownDataUpdatedAt = null;
async function pingServer() {
  try {
    // Same 6000ms budget every other network call in this app gets
    // (api(), cloudFetch(), loadCloudConfig()) -- this used to be a
    // stricter 3500ms, which meant the online/offline badge could flip
    // to "offline" purely from its own tighter timeout even while every
    // actual read/write was succeeding fine on the shared 6s budget.
    // That's a real gap on higher-latency connections (e.g. reaching a
    // cloud-hosted service from far away, or a slow mobile network).
    const res = await fetchTimeout('/api/lock-status',
      {cache:'no-store', headers: {'X-Device-Id': DEVICE_ID}}, 6000);
    if (res.ok) {
      const data = await res.json().catch(() => ({}));
      adoptCanonicalDeviceId(data && data.canonical_device_id);
      if (data && (data.kicked || data.disconnecting || data.pending_approval || data.device_limit_reached)) {
        enterBlockedState(data);
        return false;
      }
      pingMisses = 0;
      setOnline(true);
      // /api/lock-status is cheap and already polled every 6s for the
      // online badge -- data_updated_at rides along on that same request
      // instead of needing its own poll. The moment it moves past what
      // we last saw (a save on the PC, another phone, etc.), apply the
      // change right away rather than waiting for the slow 90s baseline
      // refresh below.
      const stamp = data && data.data_updated_at;
      if (stamp) {
        if (lastKnownDataUpdatedAt !== null && stamp !== lastKnownDataUpdatedAt) {
          refreshIfSafe();
        }
        lastKnownDataUpdatedAt = stamp;
      }
      return true;
    }
  } catch(e) {}
  // Unreachable, and never told us it was deliberate — treat as a
  // network/PC drop, not a disconnect: keep working from cache.
  //
  // One miss on its own isn't enough to call it "offline" -- a single
  // dropped packet or a briefly slow response every 6s is normal, not
  // a real outage, and flipping the badge on every isolated miss makes
  // it flicker even on a genuinely fine connection. Only declare
  // offline after two misses in a row (~12s of no contact), which
  // still lands well inside the 20s window the server itself uses to
  // decide a device has gone offline (see ONLINE_TIMEOUT).
  pingMisses += 1;
  if (pingMisses >= 2) {
    setOnline(false);
  }
  return false;
}
setInterval(pingServer, 6000);

// Baseline safety-net refresh -- catches anything data_updated_at-based
// detection in pingServer() might miss (e.g. this device was hidden/
// backgrounded when the change stamp ticked over) without needing its
// own tight polling interval, since real-time pickup is already handled
// above. 90s, as originally designed.
setInterval(function () {
  refreshIfSafe();
}, 90000);
window.addEventListener('online', pingServer);
window.addEventListener('offline', () => setOnline(false));

// Applies a locally-made change to the cached GET data immediately, so
// the UI reflects it right away even though the server hasn't seen it
// yet. Kept deliberately simple: it only touches the fields it's sure
// about, and otherwise just flags the record "pending sync" rather than
// guessing at recalculated totals (rent cascades, balances, etc. are
// only ever trusted once the real server has computed them).
function applyOptimisticPatch(path, method, bodyObj) {
  if (method === 'POST' && path === '/api/tenants') {
    tempIdCounter -= 1;
    const tempIdx = tempIdCounter;
    const tenant = {
      index: tempIdx, name: bodyObj.name || '(unnamed)', unit: bodyObj.unit || '',
      level: 'pending', label: 'Pending sync', days_left: null, _pending: true,
    };
    const c = loadCache();
    Object.keys(c).forEach(k => {
      if ((k === '/api/tenants' || k.indexOf('/api/tenants?') === 0) && c[k] && Array.isArray(c[k].tenants)) {
        c[k].tenants = [tenant, ...c[k].tenants];
      }
    });
    saveCache(c);
    cacheSet('/api/tenants/' + tempIdx, { tenant: {
      ...tenant, phone: bodyObj.phone||'', email: bodyObj.email||'',
      occupation: bodyObj.occupation||'', emergency_contact: bodyObj.emergency_contact||'',
      emergency_phone: bodyObj.emergency_phone||'', rent: parseFloat(bodyObj.rent)||0,
      entry_date: bodyObj.entry_date||'', notes: bodyObj.notes||'',
      deposit_paid: 0, deposit_remaining: 0,
    }});
    return tempIdx;
  }
  if (method === 'POST' && path === '/api/units') {
    const list = cacheGet('/api/units');
    if (list && Array.isArray(list.units)) {
      list.units.push({ name: bodyObj.name || '(unnamed)', rent: parseFloat(bodyObj.rent)||0, vacant: true, _pending: true });
      cacheSet('/api/units', list);
    }
    return null;
  }
  // Edits / payments / deposits / cancellations / arrears clearing on an
  // *existing* tenant or unit: don't try to recompute the financial
  // result locally — just flag it as pending so the person can see the
  // change was captured and is waiting to sync, without risking showing
  // a wrong number in the meantime.
  const tenantMatch = path.match(/^\/api\/tenants\/(-?\d+)/);
  if (tenantMatch) {
    const idx = parseInt(tenantMatch[1], 10);
    const single = cacheGet('/api/tenants/' + idx);
    if (single && single.tenant) { single.tenant._pendingSync = true; cacheSet('/api/tenants/' + idx, single); }
    const c = loadCache();
    Object.keys(c).forEach(k => {
      if ((k === '/api/tenants' || k.indexOf('/api/tenants?') === 0) && c[k] && Array.isArray(c[k].tenants)) {
        const row = c[k].tenants.find(t => t.index === idx);
        if (row) row._pendingSync = true;
      }
    });
    saveCache(c);
  }
  const unitMatch = path.match(/^\/api\/units\/([^/]+)/);
  if (unitMatch) {
    const list = cacheGet('/api/units');
    if (list && Array.isArray(list.units)) {
      const row = list.units.find(u => u.name === decodeURIComponent(unitMatch[1]));
      if (row) { row._pendingSync = true; cacheSet('/api/units', list); }
    }
  }
  return null;
}

async function flushQueue() {
  if (syncing) return;
  const q = loadQueue();
  if (!q.length) { updateSyncBadge(); return; }
  syncing = true; updateSyncBadge();
  let progressed = false;
  while (loadQueue().length) {
    const current = loadQueue();
    const op = current[0];
    try {
      const res = await fetchTimeout(op.path, {
        method: op.method, headers: {'Content-Type':'application/json'}, body: op.body,
      }, 6000);
      const data = await res.json().catch(() => ({}));
      if (!res.ok && res.status === 401) { syncing = false; updateSyncBadge(); showLock(); return; }
      // Whether the server accepted or rejected it (e.g. validation
      // error), it's been *seen* — drop it from the queue either way so
      // one bad entry can't block everything behind it forever.
      current.shift();
      saveQueue(current);
      progressed = true;
    } catch (err) {
      // Still unreachable — stop here, keep this and everything after it
      // queued in order, and try again on the next successful ping.
      syncing = false; setOnline(false); updateSyncBadge();
      return;
    }
  }
  // Fully flushed: the local cache may now be stale/wrong in ways that
  // are hard to patch precisely (server-computed totals, cascades,
  // real indexes replacing temp ones) — simplest safe move is to drop
  // the cached GET data so the next screen render pulls fresh truth.
  if (progressed) cacheDeletePrefix('/api/');
  syncing = false; updateSyncBadge();
  toast('All changes synced ✓');
  render();
}

// ── PWA: service worker + "Add to Home Screen" / install prompt ───────
// Registering the worker is what makes Chrome/Edge treat this as an
// installable app at all; the beforeinstallprompt listener captures the
// browser's native install prompt so it can be replayed from the
// Settings button instead of only appearing as a small address-bar icon.
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    // Relative registration/scope keeps this correct regardless of what
    // path the app happens to be served under. sid/key are forwarded
    // onto the SW script's own URL (readable inside it as
    // self.location.search) so the worker can precache -- and fall back
    // to -- the actual paired root instead of a bare "/", which is what
    // caused the blank white page after installing to the home screen.
    const cd = window.__CLOUD_DIRECT__;
    const swUrl = cd
      ? 'sw.js?sid=' + encodeURIComponent(cd.sessionId) + '&key=' + encodeURIComponent(cd.secretKey)
      : 'sw.js';
    navigator.serviceWorker.register(swUrl).catch(()=>{});
  });
}
let deferredInstallPrompt = null;
const isStandalone = () =>
  window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  updateInstallUI();
});
window.addEventListener('appinstalled', () => { deferredInstallPrompt = null; updateInstallUI(); });
function updateInstallUI() {
  const btn = $('#installBtn'), hint = $('#installHint');
  if (!btn || !hint) return; // Settings isn't the active page right now
  if (isStandalone()) {
    btn.style.display = 'none';
    hint.textContent = "Already installed \u2014 you're using the installed app right now.";
  } else if (deferredInstallPrompt) {
    btn.style.display = 'block';
    hint.textContent = '';
  } else {
    btn.style.display = 'none';
    const ua = navigator.userAgent || '';
    hint.textContent = /iPhone|iPad|iPod/.test(ua)
      ? 'On iPhone/iPad: tap the Share icon, then "Add to Home Screen".'
      : "Use your browser's menu (\u22ee or \u2026) and choose \\\"Install app\\\" or \\\"Add to Home Screen\\\".";
  }
}
async function triggerInstall() {
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  await deferredInstallPrompt.userChoice;
  deferredInstallPrompt = null;
  updateInstallUI();
}

// ── offline fallback shapes ──────────────────────────────────────────
// Every render*() function destructures a specific set of fields out of
// whatever api() gives it (d.tenant, d.counts.pending, d.alerts.length,
// etc). If a GET fails while offline AND there's no cached copy of that
// exact path yet (a tab that was never opened while online, or a
// tenant-detail id that was never viewed), returning one generic
// {tenants:[], units:[]} object for every endpoint left most callers
// reading a field that simply wasn't there (d.counts undefined, d.alerts
// undefined, d.tenant undefined...) which threw mid-render. That throw
// was swallowed by render()'s try/catch, so the screen was just left on
// whatever it last showed (often the "Loading…" placeholder) forever --
// this is the "offline, then reload, then nothing" bug. Shaping the
// fallback per endpoint means every render function gets the empty
// structure it actually expects instead of crashing.
function offlineDefaultFor(path) {
  const base = { error: 'offline', offline: true, no_cache: true };
  if (path.indexOf('/api/dashboard') === 0) {
    return { ...base, total_tenants: 0, total_units: 0, occupied_units: 0, vacant_units: 0,
      month_name: '', month_income: 0, full_payment_total: 0, deposit_total: 0,
      cancelled_total: 0, counts: { pending: 0, paid: 0, underpaid: 0 }, watchlist: [] };
  }
  if (/^\/api\/tenants\/[^/]+$/.test(path)) {
    return { ...base, tenant: null };
  }
  if (path.indexOf('/api/tenants') === 0) {
    return { ...base, tenants: [] };
  }
  if (path.indexOf('/api/units') === 0) {
    return { ...base, units: [] };
  }
  if (path.indexOf('/api/alerts') === 0) {
    return { ...base, alerts: [] };
  }
  if (path.indexOf('/api/history') === 0) {
    return { ...base, tenants: [] };
  }
  return { ...base, tenants: [], units: [] };
}

// ── cloud fallback ──────────────────────────────────────────────────
// When this page was loaded straight from the cloud service via the
// direct-cloud QR code (?sid=&key=), CLOUD_DIRECT is embedded before
// anything else in <head> -- see index()'s CLOUD_MODE branch in app.py
// -- and cloudCfg below is configured immediately from it, so every
// read/write always targets the cloud service, with the PC out of the
// picture entirely. If this page was instead loaded directly against a
// PC running app.py locally (e.g. for local testing) with no cloud
// pairing, /api/cloud-config tells us where a configured cloud service
// lives so reads/writes can still reach it once learned. We cache that
// response like any other GET (see cacheSet below) so it's still known
// even once the PC that served it goes away.
const CLOUD_DIRECT = window.__CLOUD_DIRECT__ || null;

let cloudCfg = CLOUD_DIRECT ? {
  configured: true,
  cloud_base_url: window.location.origin,
  session_id: CLOUD_DIRECT.sessionId,
  secret_key: CLOUD_DIRECT.secretKey,
} : null;

async function loadCloudConfig() {
  if (CLOUD_DIRECT) return; // already configured above; nothing to fetch
  const cached = cacheGet('/api/cloud-config');
  if (cached && cached.configured) cloudCfg = cached;
  try {
    const res = await fetchTimeout('/api/cloud-config', {headers:{'X-Device-Id':DEVICE_ID}}, 4000);
    if (res.ok) {
      const data = await res.json();
      cacheSet('/api/cloud-config', data);
      if (data.configured) cloudCfg = data;
    }
  } catch (e) { /* PC unreachable right now -- keep whatever was cached */ }
}

async function cloudFetch(path, opts = {}) {
  if (!cloudCfg || !cloudCfg.configured) throw new Error('cloud_not_configured');
  const headers = {
    'Content-Type': 'application/json',
    'X-Session-Id': cloudCfg.session_id,
    'X-Secret-Key': cloudCfg.secret_key,
    'X-Device-Id': DEVICE_ID,
    'X-Device-Screen': DEVICE_SCREEN_HINT,
    ...(opts.headers || {}),
  };
  const res = await fetchTimeout(cloudCfg.cloud_base_url + path, { ...opts, headers }, 6000);
  const data = await res.json().catch(() => ({}));
  if (!res.ok) { data.ok = false; data.error = data.error || 'cloud_request_failed'; }
  return data;
}

// ── api helper ───────────────────────────────────────────────────────
async function api(path, opts={}) {
  const method = (opts.method || 'GET').toUpperCase();

  if (method === 'GET') {
    if (/^\/api\/tenants\/-\d+/.test(path)) {
      // Temp/offline-created record — never exists on the server, so
      // don't waste time trying to reach it.
      const cached = cacheGet(path);
      if (cached !== undefined) return cached;
      return { error: 'not_found' };
    }

    // ── Cloud-first ──────────────────────────────────────────────────
    // Reads are served from the cloud database whenever it's reachable,
    // full stop. When this page was loaded via the direct-cloud QR code,
    // that's the only path that's ever configured, so it's also the
    // only place reads ever come from. Every successful read here
    // refreshes the local cache too, so the app still has something to
    // show even with no connectivity at all.
    //
    // loadCloudConfig() otherwise only ever ran once, at boot() -- if
    // that single attempt missed (briefly unreachable at that exact
    // moment), cloudCfg stayed unconfigured for the whole session.
    // Retrying it here, every time it's missing, means a later
    // reconnect actually gets picked up instead of needing a reload.
    if (!cloudCfg || !cloudCfg.configured) {
      await loadCloudConfig();
    }
    if (cloudCfg && cloudCfg.configured) {
      try {
        const cloudData = await cloudFetch(path);
        if (cloudData && cloudData.ok !== false) {
          cacheSet(path, cloudData);
          return cloudData;
        }
      } catch (e) { /* cloud unreachable right now -- fall back below */ }
    }

    // ── Local/PC fallback ────────────────────────────────────────────
    // Only reached when the cloud isn't configured yet or isn't
    // reachable right now -- e.g. this page is running directly against
    // a local app.py instance with no cloud pairing set up. Still
    // responsible for the isOnline flag -- that badge specifically
    // tracks "is the server this page came from reachable", which is
    // exactly what pingServer() also polls independently every 6s.
    try {
      const res = await fetchTimeout(path, {headers:{'Content-Type':'application/json','X-Device-Id':DEVICE_ID}, ...opts}, 4000);
      if (res.status === 401) { showLock(); throw new Error('locked'); }
      if (res.status === 403) {
        const data403 = await res.json().catch(()=>({}));
        if (data403 && data403.error === 'kicked') {
          handleKicked(); throw new Error('kicked');
        }
        if (data403 && data403.error === 'pending_approval') {
          handlePendingApproval(); throw new Error('pending_approval');
        }
        if (data403 && data403.error === 'device_limit_reached') {
          showDeviceLimit(data403.max_devices); throw new Error('device_limit');
        }
      }
      if (res.status === 502 || res.status === 503 || res.status === 504) {
        // A genuine gateway/server error, not a real answer from the
        // app. Treat exactly like a network failure below: fall back to
        // cache, don't let this overwrite the real cached data or get
        // reported as "Online".
        throw new Error('gateway_unreachable');
      }
      const data = await res.json().catch(()=>({}));
      if (!res.ok && data.ok !== false) { data.error = data.error || 'Request failed.'; }
      cacheSet(path, data);
      setOnline(true);
      return data;
    } catch (err) {
      if (err && (err.message === 'locked' || err.message === 'device_limit' || err.message === 'kicked' || err.message === 'pending_approval')) throw err;
      setOnline(false);
      const cached = cacheGet(path);
      if (cached !== undefined) return cached;
      return offlineDefaultFor(path);
    }
  }

  // Mutating request (POST/PUT/DELETE)
  try {
    const res = await fetchTimeout(path, {headers:{'Content-Type':'application/json','X-Device-Id':DEVICE_ID}, ...opts}, 6000);
    if (res.status === 401) { showLock(); throw new Error('locked'); }
    if (res.status === 403) {
      const data403 = await res.json().catch(()=>({}));
      if (data403 && data403.error === 'kicked') {
        handleKicked(); throw new Error('kicked');
      }
      if (data403 && data403.error === 'pending_approval') {
        handlePendingApproval(); throw new Error('pending_approval');
      }
      if (data403 && data403.error === 'device_limit_reached') {
        showDeviceLimit(data403.max_devices); throw new Error('device_limit');
      }
    }
    if (res.status === 502 || res.status === 503 || res.status === 504) {
      // A genuine gateway/server error, not a real answer -- same
      // situation as a network failure, so fall into the catch block
      // below and queue this change to sync once reconnected.
      throw new Error('gateway_unreachable');
    }
    const data = await res.json().catch(()=>({}));
    if (!res.ok && data.ok !== false) { data.error = data.error || 'Request failed.'; }
    setOnline(true);
    return data;
  } catch (err) {
    if (err && (err.message === 'locked' || err.message === 'device_limit' || err.message === 'kicked' || err.message === 'pending_approval')) throw err;
    setOnline(false);
    if (!cloudCfg || !cloudCfg.configured) {
      await loadCloudConfig();
    }
    if (cloudCfg && cloudCfg.configured) {
      try {
        const cloudData = await cloudFetch(path, { method, body: opts.body });
        if (cloudData && cloudData.ok !== false) {
          toast('Saved directly to the cloud — PC will sync when it reconnects.');
          return cloudData;
        }
      } catch (e2) { /* cloud unreachable too -- fall through to local queue */ }
    }
    let bodyObj = {};
    try { bodyObj = opts.body ? JSON.parse(opts.body) : {}; } catch(e) {}
    const tempIdx = applyOptimisticPatch(path, method, bodyObj);
    const q = loadQueue();
    q.push({ path, method, body: opts.body || null, ts: Date.now() });
    saveQueue(q);
    toast('Offline — change saved, will sync once reconnected.');
    return { ok: true, offline: true, queued: true, index: tempIdx };
  }
}
function toast(msg) {
  const el = document.createElement('div');
  el.className = 'toast';
  el.textContent = msg;
  $('#toastRoot').appendChild(el);
  setTimeout(()=>el.remove(), 2600);
}

// ── lock screen ──────────────────────────────────────────────────────
let pinBuffer = '';
function buildKeypad() {
  const kp = $('#keypad'); kp.innerHTML = '';
  const keys = ['1','2','3','4','5','6','7','8','9','','0','⌫'];
  keys.forEach(k => {
    const b = document.createElement('button');
    b.textContent = k;
    if (!k) { b.style.visibility='hidden'; }
    else b.onclick = () => onKey(k);
    kp.appendChild(b);
  });
}
function renderDots() {
  const dots = $('#pinDots'); dots.innerHTML = '';
  for (let i=0;i<Math.max(4,pinBuffer.length);i++) {
    const d = document.createElement('div');
    d.className = 'd' + (i < pinBuffer.length ? ' filled':'');
    dots.appendChild(d);
  }
}
async function onKey(k) {
  $('#lockErr').textContent = '';
  if (k === '⌫') { pinBuffer = pinBuffer.slice(0,-1); renderDots(); return; }
  if (pinBuffer.length >= 8) return;
  pinBuffer += k;
  renderDots();
  if (pinBuffer.length >= 4) {
    try {
      const res = await fetchTimeout('/api/unlock', {method:'POST', headers:{'Content-Type':'application/json','X-Device-Id':DEVICE_ID}, body: JSON.stringify({pin: pinBuffer})}, 4000);
      const d = await res.json();
      if (d.ok) { pinBuffer=''; setOnline(true); hideLock(); boot(); }
      else if (d.error === 'kicked') {
        handleKicked(); pinBuffer=''; renderDots();
      }
      else if (d.error === 'pending_approval') {
        handlePendingApproval(); pinBuffer=''; renderDots();
      }
      else if (d.error === 'device_limit_reached') {
        showDeviceLimit(d.max_devices); pinBuffer=''; renderDots();
      }
      else if (pinBuffer.length >= 8 || d.error) {
        $('#lockErr').textContent = d.error || 'Incorrect PIN.';
        pinBuffer=''; renderDots();
      }
    } catch (err) {
      setOnline(false);
      $('#lockErr').textContent = "Can't reach the PC to verify your PIN right now — try again once it's back online.";
      pinBuffer=''; renderDots();
    }
  }
}
function showLock() {
  $('#lockscreen').style.display = 'flex';
  $('#app').style.display = 'none';
  $('#tabbar').style.display = 'none';
  pinBuffer = ''; renderDots();
  buildKeypad();
}
function hideLock() {
  $('#lockscreen').style.display = 'none';
  $('#app').style.display = 'block';
  $('#tabbar').style.display = 'flex';
}
function lockNow() {
  fetchTimeout('/api/lock', {method:'POST'}, 3000).catch(()=>{}).then(()=>showLock());
}

// ── modal helper ─────────────────────────────────────────────────────
function openModal(html, opts) {
  const dismissible = !opts || opts.dismissible !== false;
  const root = $('#modalRoot');
  root.innerHTML = `<div class="modal-backdrop" ${dismissible ? `onclick="if(event.target===this) closeModal()"` : ''}>
    <div class="modal" style="position:relative;">
      ${dismissible ? `<button class="icon-btn close-x" style="background:var(--teal-soft);color:var(--teal-deep);" onclick="closeModal()">✕</button>` : ''}
      ${html}
    </div>
  </div>`;
}
function closeModal() { $('#modalRoot').innerHTML = ''; }

// ── double-submit guard ──────────────────────────────────────────────
// Fixes the "tapped Save/Confirm twice because it felt slow" bug: a
// second tap before the first request finished used to re-run the same
// handler and duplicate the change (double payment, double edit,
// double cancellation, etc). This intercepts every click on any
// <button> that still has its inline onclick -- covers .btn (Save,
// Confirm, Record Payment...) as well as icon-buttons like the ↺
// cancel-transaction control and the lock button -- disables the
// button the instant the tap lands (before the handler even starts, so
// there's no window for a second tap to slip through), runs the
// handler, then re-enables the button only once that handler --
// including any awaited network call -- has actually settled. Buttons
// wired up purely with addEventListener (tabs, filter pills) have no
// inline onclick attribute and so are untouched.
document.addEventListener('click', function (e) {
  const btn = e.target.closest('button[onclick]');
  if (!btn || btn.disabled) return;
  const code = btn.getAttribute('onclick');
  if (!code) return;
  e.preventDefault();
  e.stopImmediatePropagation();
  btn.disabled = true;
  Promise.resolve()
    .then(() => new Function(code).call(btn))
    .catch(err => console.error(err))
    .finally(() => { if (document.body.contains(btn)) btn.disabled = false; });
}, true);

// ── tabs ─────────────────────────────────────────────────────────────
$$('.tab').forEach(t => t.addEventListener('click', () => switchTab(t.dataset.tab)));
function switchTab(tab) {
  state.tab = tab;
  const moreGroup = ['more', 'history', 'settings'];
  $$('.tab').forEach(t => {
    const dt = t.dataset.tab;
    const active = (dt === tab) || (dt === 'more' && moreGroup.includes(tab));
    t.classList.toggle('active', active);
  });
  render();
}

async function refreshAlertBadge() {
  try {
    const d = await api('/api/dashboard');
    const n = d.watchlist ? d.watchlist.length : 0;
    const dot = $('#alertDot');
    if (n>0) { dot.style.display='flex'; dot.textContent = n>9?'9+':n; }
    else dot.style.display='none';
  } catch(e) {}
}

// ── render router ────────────────────────────────────────────────────
async function render() {
  const main = $('#main');
  // Don't blank the screen immediately -- on a fast LAN the fetch below
  // usually resolves in well under 100ms, so clearing main first just
  // produces a visible flash/flicker on every tab-bar tap for no benefit.
  // Only show the Loading placeholder if this render is actually slow.
  let loadingShown = false;
  const loadingTimer = setTimeout(() => {
    loadingShown = true;
    main.innerHTML = '<div class="empty">Loading…</div>';
  }, 220);
  try {
    if (state.tab === 'dashboard') await renderDashboard();
    else if (state.tab === 'tenants') await renderTenants();
    else if (state.tab === 'units') await renderUnits();
    else if (state.tab === 'alerts') await renderAlerts();
    else if (state.tab === 'more') renderMoreMenu();
    else if (state.tab === 'history') await renderHistory();
    else if (state.tab === 'settings') await renderSettings();
    else if (state.tab === 'tenant-detail') await renderTenantDetail(state.selectedIdx);
    else if (state.tab === 'add-tenant') renderAddTenant();
  } catch(e) {
    // Previously this just logged and returned, leaving whatever was on
    // screen (often the "Loading…" placeholder from above) stuck there
    // forever with no way out. Now it always leaves the person with a
    // visible, actionable screen instead of a silent dead end.
    console.error(e);
    main.innerHTML = `<div class="empty">
      <div class="big">${isOnline ? '⚠️' : '📴'}</div>
      ${isOnline ? "Something went wrong loading this." : "You're offline and this hasn't loaded before, so there's nothing cached to show yet."}
      <div style="margin-top:14px;"><button class="btn btn-primary" onclick="render()">Try Again</button></div>
    </div>`;
  }
  clearTimeout(loadingTimer);
  refreshAlertBadge();
}

// ── DASHBOARD (Home) ───────────────────────────────────────────────────
// Mirrors the desktop app's Dashboard tab: four colored stat cards in a
// 2×2 grid, each with a clickable footer link that jumps to the matching
// section — same accent colors, same card contents, nothing else on the
// page (no Add Tenant button, no extra list below).
const D_BLUE   = '#1A73E8';   // Total Tenants
const D_AMBER  = '#F9A825';   // Total Units
const D_GREEN  = '#2E7D32';   // Total Income
const D_ORANGE = '#E65100';   // Rent Alerts

function dcard({accent, icon, title, value, subs, actionLabel, actionTab}) {
  const subRow = subs ? `<div class="dcard-subrow">${
    subs.map(([label, val, color]) => `
      <div class="dcard-sub">
        <div class="v" style="color:${color}">${val}</div>
        <div class="l">${label}</div>
      </div>`).join('')
  }</div>` : '';
  return `<div class="dcard">
    <div class="dcard-top" style="background:${accent}">
      <span class="dcard-icon">${icon}</span>
      <div class="dcard-title">${title}</div>
      <div class="dcard-value">${value}</div>
    </div>
    <div class="dcard-footer">
      ${subRow}
      <button class="dcard-action" style="color:${accent}" onclick="switchTab('${actionTab}')">${actionLabel} ›</button>
    </div>
  </div>`;
}

async function renderDashboard() {
  const d = await api('/api/dashboard');
  if (d.no_cache) {
    // Offline and the dashboard was never loaded on this device before --
    // showing "0 tenants · 0 units" here would look like real data instead
    // of "we don't know yet", so say that plainly instead.
    $('#headerSub').textContent = '';
    $('#main').innerHTML = `<div class="empty"><div class="big">📴</div>You're offline, and this device hasn't loaded any data yet. Connect to the same Wi-Fi as the PC (or check its internet), then try again.</div>`;
    return;
  }
  $('#headerSub').textContent = `${d.total_tenants} tenants · ${d.total_units} units`;

  $('#main').innerHTML = `
    <div class="dgrid">
      ${dcard({
        accent: D_BLUE, icon:'🧑', title:'Total Tenants', value: d.total_tenants,
        actionLabel:'View Tenants', actionTab:'tenants'
      })}
      ${dcard({
        accent: D_AMBER, icon:'🏠', title:'Total Units', value: d.total_units,
        subs: [['Occupied', d.occupied_units, D_BLUE], ['Vacant', d.vacant_units, 'var(--muted)']],
        actionLabel:'View Units', actionTab:'units'
      })}
      ${dcard({
        accent: D_GREEN, icon:'💰', title:`Total Income — ${d.month_name}`,
        value: fmt(d.month_income),
        subs: [['Full', fmt(d.full_payment_total), 'var(--good)'],
               ['Deposits', fmt(d.deposit_total), 'var(--teal-deep)'],
               ['Cancelled', fmt(d.cancelled_total), 'var(--danger)']],
        actionLabel:'View Records', actionTab:'history'
      })}
      ${dcard({
        accent: D_ORANGE, icon:'🔔', title:'Rent Alerts', value: d.counts.pending,
        subs: [['Pending', d.counts.pending, 'var(--warn)'],
               ['Paid in Full', d.counts.paid, 'var(--good)'],
               ['Installments', d.counts.underpaid, 'var(--teal-deep)']],
        actionLabel:'View Alerts', actionTab:'alerts'
      })}
    </div>
  `;
}

function tenantRowHtml(t) {
  const chipClass = t.level === 'paid' ? 'chip-paid' : (t.level==='underpaid'?'chip-underpaid':'chip-pending');
  const daysTxt = t.days_left===null||t.days_left===undefined ? '' :
    (t.days_left<0 ? `Overdue ${Math.abs(t.days_left)}d` : `${t.days_left}d left`);
  const initials = (t.name||'?').split(' ').filter(Boolean).slice(0,2).map(w=>w[0]).join('').toUpperCase();
  const chip = (t._pending || t._pendingSync)
    ? `<span class="chip chip-underpaid">🕓 Pending sync</span>`
    : `<span class="chip ${chipClass}">${t.label}</span>`;
  return `<div class="tenant-row" onclick="openTenant(${t.index})">
    <div class="avatar">${initials||'?'}</div>
    <div class="meta">
      <div class="name">${escapeHtml(t.name)}</div>
      <div class="sub">${escapeHtml(t.unit)} · ${daysTxt}</div>
    </div>
    ${chip}
  </div>`;
}
function escapeHtml(s){ const d=document.createElement('div'); d.textContent=s??''; return d.innerHTML; }

const MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function fmtPeriodDate(d) {
  return `${d.getDate()} ${MONTH_ABBR[d.getMonth()]}, ${String(d.getFullYear()).slice(-2)}`;
}
function abbrevPeriod(fromIso, toIso) {
  // Must always agree with the record's actual from/to fields (as shown
  // in the desktop app's Records export), so this renders the exact
  // dates rather than collapsing them to a month/year label -- rounding
  // to whole months previously hid the real (often mid-month) cycle
  // boundaries and could visibly disagree with the underlying dates.
  const f = _parseIsoDate(fromIso), t = _parseIsoDate(toIso);
  if (!f || !t) return `${fromIso || '—'} to ${toIso || '—'}`;
  return `${fmtPeriodDate(f)} to ${fmtPeriodDate(t)}`;
}
function _parseIsoDate(s) {
  if (!s || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return null;
  const [y, m, d] = s.split('-').map(Number);
  return new Date(y, m - 1, d);
}

function openTenant(idx) { state.selectedIdx = idx; state.tab = 'tenant-detail'; render(); }

// ── TENANTS LIST ─────────────────────────────────────────────────────
async function renderTenants() {
  const params = new URLSearchParams({q: state.q, filter: state.filter});
  const d = await api('/api/tenants?' + params.toString());
  $('#headerSub').textContent = `${d.tenants.length} shown`;
  const filters = [['all','All'],['paid','Paid'],['underpaid','Instalments'],['pending','Pending']];
  const emptyMsg = d.no_cache
    ? `<div class="empty"><div class="big">📴</div>You're offline and this hasn't loaded before, so there's nothing cached to show yet.</div>`
    : `<div class="empty"><div class="big">👤</div>No tenants match.</div>`;
  const rows = d.tenants.map(tenantRowHtml).join('') || emptyMsg;
  $('#main').innerHTML = `
    <div class="searchbar"><span>🔎</span><input id="searchInput" placeholder="Search name or unit" value="${escapeHtml(state.q)}"></div>
    <div class="filters">${filters.map(([k,l])=>`<div class="filter-pill ${state.filter===k?'active':''}" data-f="${k}">${l}</div>`).join('')}</div>
    <button class="btn btn-primary btn-full" style="margin-bottom:14px;" onclick="state.tab='add-tenant'; render();">＋ Add Tenant</button>
    <div class="card" style="padding:4px 12px;">${rows}</div>
  `;
  $('#searchInput').addEventListener('input', debounce(e => { state.q = e.target.value; renderTenants(); }, 300));
  $$('.filter-pill').forEach(p => p.addEventListener('click', () => { state.filter = p.dataset.f; renderTenants(); }));
}
function debounce(fn, ms) { let t; return (...a)=>{ clearTimeout(t); t=setTimeout(()=>fn(...a), ms); }; }

// ── ADD TENANT ───────────────────────────────────────────────────────
async function renderAddTenant() {
  $('#headerSub').textContent = 'New tenant';
  const u = await api('/api/units/vacant');
  const opts = u.units.map(x => `<option value="${escapeHtml(x.name)}" data-rent="${x.rent}">${escapeHtml(x.name)} — ${fmt(x.rent)}/mo</option>`).join('');
  $('#main').innerHTML = `
    <button class="btn btn-ghost" style="margin-bottom:12px;" onclick="state.tab='tenants'; render();">← Back</button>
    <div class="card">
      <h2 style="margin-top:0;">Add Tenant</h2>
      <label class="field">Full Name *</label><input id="f_name">
      <label class="field">Unit *</label><select id="f_unit"><option value="">Select a vacant unit…</option>${opts}</select>
      <label class="field">Phone *</label><input id="f_phone">
      <div class="row"><div><label class="field">Email</label><input id="f_email"></div>
      <div><label class="field">Occupation</label><input id="f_occupation"></div></div>
      <div class="row"><div><label class="field">Emergency Contact</label><input id="f_emergency_contact"></div>
      <div><label class="field">Emergency Phone</label><input id="f_emergency_phone"></div></div>
      <label class="field">Monthly Rent (UGX)</label><input id="f_rent" inputmode="numeric">
      <label class="field">Move-in Date *</label><input id="f_entry_date" type="date" value="${todayStr()}">
      <label class="field">Notes</label><textarea id="f_notes" rows="2"></textarea>
      <div class="err" id="addErr"></div>
      <button class="btn btn-primary btn-full" style="margin-top:12px;" onclick="submitAddTenant()">Save Tenant</button>
    </div>
  `;
  $('#f_unit').addEventListener('change', e => {
    const opt = e.target.selectedOptions[0];
    if (opt && opt.dataset.rent) $('#f_rent').value = opt.dataset.rent;
  });
}
async function submitAddTenant(replace=false) {
  const body = {
    name: $('#f_name').value, unit: $('#f_unit').value, phone: $('#f_phone').value,
    email: $('#f_email').value, occupation: $('#f_occupation').value,
    emergency_contact: $('#f_emergency_contact').value, emergency_phone: $('#f_emergency_phone').value,
    rent: $('#f_rent').value, entry_date: $('#f_entry_date').value, notes: $('#f_notes').value, replace,
  };
  const d = await api('/api/tenants', {method:'POST', body: JSON.stringify(body)});
  if (d.ok) { toast('Tenant saved.'); state.tab='tenants'; render(); return; }
  if (d.error === 'unit_taken') {
    if (confirm(d.message + ' Replace with this new tenant?')) return submitAddTenant(true);
    return;
  }
  $('#addErr').textContent = d.error || 'Could not save tenant.';
}

// ── TENANT DETAIL ────────────────────────────────────────────────────
async function renderTenantDetail(idx) {
  const d = await api('/api/tenants/' + idx);
  const t = d.tenant;
  if (!t) {
    // Offline, and this specific tenant was never viewed/cached before —
    // nothing safe to render, so say so plainly instead of crashing on
    // t.unit below.
    $('#headerSub').textContent = 'Tenant';
    $('#main').innerHTML = `
      <button class="btn btn-ghost" style="margin-bottom:12px;" onclick="state.tab='tenants'; render();">← Back</button>
      <div class="empty"><div class="big">📴</div>You're offline and this tenant hasn't loaded before, so there's nothing cached to show yet.</div>`;
    return;
  }
  $('#headerSub').textContent = t.unit;

  if (t._pending) {
    // Created while offline — has no real record on the server yet, so
    // there's nothing to safely edit/pay against until it syncs.
    $('#main').innerHTML = `
      <button class="btn btn-ghost" style="margin-bottom:12px;" onclick="state.tab='tenants'; render();">← Back</button>
      <div class="pending-note">🕓 This tenant was added while offline and hasn't synced to the PC yet. It'll sync automatically once reconnected — editing and payments open up after that.</div>
      <div class="card">
        <h2 style="margin-top:0;">${escapeHtml(t.name)}</h2>
        <div style="font-size:13.5px;color:var(--muted);">${escapeHtml(t.unit)} · ${fmt(t.rent)}/mo</div>
        <hr class="sep">
        <div style="font-size:13px;"><b>Phone:</b> ${escapeHtml(t.phone||'—')}</div>
        <div style="font-size:13px;margin-top:4px;"><b>Move-in:</b> ${escapeHtml(t.entry_date||'—')}</div>
      </div>`;
    return;
  }
  const pendingBanner = t._pendingSync
    ? `<div class="pending-note">🕓 A change to this tenant is queued and will sync as soon as this PC is reachable again. Figures below are from the last sync.</div>`
    : '';
  const chipClass = t.level === 'paid' ? 'chip-paid' : (t.level==='underpaid'?'chip-underpaid':'chip-pending');

  const depPct = t.rent>0 ? Math.min(100, Math.round((t.deposit_paid/t.rent)*100)) : 0;
  const depBlock = (t.level==='underpaid') ? `
    <div class="card">
      <div class="section-title" style="margin-top:0;">Instalment Progress</div>
      <div style="display:flex;justify-content:space-between;font-size:13px;"><span>${fmt(t.deposit_paid)} paid</span><span>${fmt(t.deposit_remaining)} left</span></div>
      <div class="progress-bar"><div style="width:${depPct}%"></div></div>
    </div>` : '';

  const arrearsBlock = t.rent_increase_due > 0 ? `
    <div class="card" style="border-color:#FFE3B0;background:#FFFBF0;">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <div><b>Rent Increase Arrears</b><div class="sub" style="color:var(--muted);font-size:12px;">${fmt(t.rent_increase_due)} owed from a mid-cycle increase</div></div>
        <button class="btn btn-ghost" onclick="openClearArrears(${idx}, ${t.rent_increase_due})">Clear</button>
      </div>
    </div>` : '';

  const payHist = t.payment_history.map((r,i)=>txnRow(t, r, 'payment_history', origIdx(t.payment_history,i))).join('');
  const depHist = t.deposit_history.map((r,i)=>txnRow(t, r, 'deposit_history', origIdx(t.deposit_history,i))).join('');

  $('#main').innerHTML = `
    <button class="btn btn-ghost" style="margin-bottom:12px;" onclick="state.tab='tenants'; render();">← Back</button>
    ${pendingBanner}
    <div class="card">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;">
        <div>
          <h2 style="margin:0 0 2px;">${escapeHtml(t.name)}</h2>
          <div class="sub" style="color:var(--muted);font-size:13px;">${escapeHtml(t.unit)} · ${fmt(t.rent)}/mo</div>
        </div>
        <span class="chip ${chipClass}">${t.label}</span>
      </div>
      <hr class="sep">
      <div style="font-size:13px;line-height:1.9;">
        <div>📞 ${escapeHtml(t.phone||'—')}</div>
        <div>✉️ ${escapeHtml(t.email||'—')}</div>
        <div>💼 ${escapeHtml(t.occupation||'—')}</div>
        <div>🏁 Move-in: ${escapeHtml(t.entry_date||'—')}</div>
        <div>📅 Due date: ${escapeHtml(t.due_date||'Not yet set')}</div>
      </div>
      <div class="row" style="margin-top:10px;">
        <button class="btn btn-ghost" onclick="openEditTenant(${idx})">✎ Edit Tenant</button>
        <button class="btn btn-danger" onclick="deleteTenant(${idx})">🗑 Delete Tenant</button>
      </div>
    </div>

    ${leaseProgressBlock(t)}
    ${depBlock}
    ${arrearsBlock}

    <div class="card">
      <div class="row">
        <button class="btn btn-primary" onclick="openPaymentModal(${idx})">💳 Pay Rent</button>
      </div>
      <div class="row" style="margin-top:10px;">
        <button class="btn btn-ghost" onclick="openDepositModal(${idx})">＋ Record Instalment</button>
      </div>
    </div>

    <div class="section-title">Payment History</div>
    <div class="card" style="padding:4px 14px;">${payHist || '<div class="empty" style="padding:16px;">No full payments yet.</div>'}</div>

    <div class="section-title">Instalment / Deposit History</div>
    <div class="card" style="padding:4px 14px;">${depHist || '<div class="empty" style="padding:16px;">No instalments yet.</div>'}</div>
  `;
}
// history arrays are returned reversed (most-recent-first); recover original index for cancel calls
function origIdx(arr, displayI) { return arr.length - 1 - displayI; }

function txnRow(t, r, key, origI) {
  const cancelled = r._cancelled;
  const label = key==='payment_history' ? 'Full Payment' : 'Deposit';
  const periodTxt = (r.from_date && r.to_date) ? abbrevPeriod(r.from_date, r.to_date) : '';
  const monthsTxt = r.months ? `${r.months} month${r.months==1?'':'s'}` : '';
  const subTxt = [escapeHtml(r.date||''), monthsTxt, periodTxt].filter(Boolean).join(' · ');
  return `<div class="txn-item">
    <div>
      <div style="font-size:13px;">${label}${cancelled?' <span style="color:var(--danger);font-weight:600;">(Cancelled)</span>':''}</div>
      <div class="sub" style="font-size:11.5px;color:var(--muted);">${subTxt}</div>
    </div>
    <div style="display:flex;align-items:center;gap:10px;">
      <span class="amt ${cancelled?'cancelled':''}">${fmt(r.amount)}</span>
      ${!cancelled ? `<button class="icon-btn" style="width:30px;height:30px;background:var(--teal-soft);color:var(--teal-deep);font-size:13px;" onclick="cancelTxn(${t.index}, '${key}', ${origI})">↺</button>`:''}
    </div>
  </div>`;
}

async function cancelTxn(idx, key, recIdx) {
  if (!confirm("Cancel this record and reverse its effect on the tenant's account?")) return;
  const d = await api(`/api/tenants/${idx}/cancel`, {method:'POST', body: JSON.stringify({history_key:key, record_index:recIdx})});
  if (d.ok) { toast(`Reversed ${fmt(d.result.total_amount)}.`); state.selectedIdx=idx; state.tab='tenant-detail'; render(); }
  else toast(d.error || 'Could not cancel.');
}

function openEditTenant(idx) {
  api('/api/tenants/'+idx).then(d => {
    const t = d.tenant;
    openModal(`
      <h2>Edit Tenant</h2>
      <label class="field">Full Name</label><input id="e_name" value="${escapeHtml(t.name)}">
      <div class="row"><div><label class="field">Phone</label><input id="e_phone" value="${escapeHtml(t.phone)}"></div>
      <div><label class="field">Email</label><input id="e_email" value="${escapeHtml(t.email)}"></div></div>
      <label class="field">Occupation</label><input id="e_occupation" value="${escapeHtml(t.occupation)}">
      <div class="row"><div><label class="field">Emergency Contact</label><input id="e_emergency_contact" value="${escapeHtml(t.emergency_contact)}"></div>
      <div><label class="field">Emergency Phone</label><input id="e_emergency_phone" value="${escapeHtml(t.emergency_phone)}"></div></div>
      <div class="row"><div><label class="field">Move-in Date</label><input id="e_entry_date" type="date" value="${t.entry_date}"></div>
      <div><label class="field">Due Date</label><input id="e_due_date" type="date" value="${t.due_date}"></div></div>
      <label class="field">Rent (UGX)</label><input id="e_rent" value="${t.rent}">
      <label class="field">Notes</label><textarea id="e_notes" rows="2">${escapeHtml(t.notes)}</textarea>
      <div class="err" id="editErr"></div>
      <button class="btn btn-primary btn-full" style="margin-top:12px;" onclick="submitEditTenant(${idx})">Save Changes</button>
      <button class="btn btn-danger btn-full" style="margin-top:10px;" onclick="deleteTenant(${idx})">Delete Tenant</button>
    `);
  });
}
async function submitEditTenant(idx) {
  const body = {
    name: $('#e_name').value, phone: $('#e_phone').value, email: $('#e_email').value,
    occupation: $('#e_occupation').value, emergency_contact: $('#e_emergency_contact').value,
    emergency_phone: $('#e_emergency_phone').value, entry_date: $('#e_entry_date').value,
    due_date: $('#e_due_date').value, rent: $('#e_rent').value, notes: $('#e_notes').value,
  };
  const d = await api('/api/tenants/'+idx, {method:'PUT', body: JSON.stringify(body)});
  if (d.ok) { closeModal(); toast('Tenant updated.'); state.selectedIdx=idx; state.tab='tenant-detail'; render(); }
  else $('#editErr').textContent = d.error || 'Could not save.';
}
async function deleteTenant(idx) {
  if (!confirm('Permanently delete this tenant? This cannot be undone.')) return;
  const d = await api('/api/tenants/'+idx, {method:'DELETE'});
  if (d.ok) { closeModal(); toast('Tenant deleted.'); state.tab='tenants'; render(); }
}

// ── Month picker (shared by Pay Rent + Record Instalment) ──────────────
// Months are always ticked contiguously starting from the tenant's first
// still-open month: ticking a row ticks it and every open row before it;
// unticking a row unticks it and every open row after it. Already-paid
// (cleared) months are shown for context but can't be selected again.
let _mp = { open: [], cleared: [], selected: 0, pending: 0 };

async function loadMonthPicker(idx) {
  const d = await api(`/api/tenants/${idx}/months`);
  _mp = { open: d.open || [], cleared: d.cleared || [], selected: 0, pending: 0 };
}

function monthPickerHtml() {
  return `
    <label class="field">Month(s)</label>
    <div class="month-picker-control" id="mp_control" onclick="toggleMonthPickerPanel()">
      <span id="mp_summary" style="color:var(--muted);">Select month(s)</span>
      <span class="mp-caret">▾</span>
    </div>
    <div class="month-picker-panel" id="mp_panel" style="display:none;">
      <div class="month-picker-list" id="mp_list"></div>
      <div class="month-picker-actions">
        <button class="btn btn-ghost" type="button" style="flex:1;" onclick="cancelMonthPicker()">Cancel</button>
        <button class="btn btn-primary" type="button" style="flex:1;" onclick="confirmMonthPicker()">OK</button>
      </div>
    </div>`;
}

function renderMonthPickerList() {
  const list = $('#mp_list');
  if (!list) return;
  const clearedHtml = _mp.cleared.map(m => `
    <div class="month-row cleared">
      <input type="checkbox" checked disabled>
      <span class="month-label">${escapeHtml(m.label)}</span>
      <span class="month-tag">Paid</span>
    </div>`).join('');
  const openHtml = _mp.open.map((m, i) => `
    <div class="month-row open ${i < _mp.pending ? 'ticked' : ''}" onclick="toggleMonthRow(${i})">
      <input type="checkbox" ${i < _mp.pending ? 'checked' : ''}>
      <span class="month-label">${escapeHtml(m.label)}</span>
    </div>`).join('');
  list.innerHTML = clearedHtml + openHtml;
}

function toggleMonthRow(i) {
  // Ticking row i ticks it and every open month before it (keeps the
  // selection contiguous from the first open month); unticking it drops
  // it and every open month after it, for the same reason.
  _mp.pending = (i < _mp.pending) ? i : i + 1;
  renderMonthPickerList();
}

function toggleMonthPickerPanel() {
  const panel = $('#mp_panel');
  if (panel.style.display !== 'none') { panel.style.display = 'none'; return; }
  _mp.pending = _mp.selected;
  renderMonthPickerList();
  panel.style.display = 'block';
}

function cancelMonthPicker() {
  _mp.pending = _mp.selected;
  $('#mp_panel').style.display = 'none';
}

function confirmMonthPicker() {
  _mp.selected = _mp.pending;
  $('#mp_panel').style.display = 'none';
  updateMonthPickerSummary();
}

function updateMonthPickerSummary() {
  const el = $('#mp_summary');
  if (!el) return;
  if (_mp.selected <= 0) {
    el.textContent = 'Select month(s)';
    el.style.color = 'var(--muted)';
  } else {
    const first = _mp.open[0], last = _mp.open[_mp.selected - 1];
    el.textContent = abbrevPeriod(first.from, last.to);
    el.style.color = 'var(--ink)';
  }
  const btn = document.getElementById('txnSubmitBtn');
  if (btn) btn.disabled = _mp.selected <= 0;
}

async function openPaymentModal(idx) {
  await loadMonthPicker(idx);
  openModal(`
    <h2>Record Payment</h2>
    <div class="desc">Select the month(s) being paid in full — the due date moves forward by however many months you pick.</div>
    ${monthPickerHtml()}
    <div class="err" id="payErr"></div>
    <button class="btn btn-primary btn-full" style="margin-top:12px;" id="txnSubmitBtn" disabled onclick="submitPayment(${idx})">✓ Record Payment</button>
  `);
  renderMonthPickerList();
}
async function submitPayment(idx) {
  if (_mp.selected <= 0) { $('#payErr').textContent = 'Select at least one month.'; return; }
  const d = await api(`/api/tenants/${idx}/payment`, {method:'POST', body: JSON.stringify({months: _mp.selected})});
  if (d.ok) { closeModal(); toast(`${fmt(d.result.amount)} recorded. New due date: ${d.result.due_date}`); state.selectedIdx=idx; state.tab='tenant-detail'; render(); }
  else $('#payErr').textContent = d.error || 'Could not record payment.';
}

async function openDepositModal(idx) {
  await loadMonthPicker(idx);
  openModal(`
    <h2>Record Instalment / Deposit</h2>
    <div class="desc">Partial payments accumulate toward the selected month(s)' rent.</div>
    ${monthPickerHtml()}
    <label class="field">Amount (UGX)</label><input id="d_amount" inputmode="numeric">
    <div class="err" id="depErr"></div>
    <button class="btn btn-primary btn-full" style="margin-top:12px;" id="txnSubmitBtn" disabled onclick="submitDeposit(${idx})">＋ Record Instalment</button>
  `);
  renderMonthPickerList();
}
async function submitDeposit(idx) {
  if (_mp.selected <= 0) { $('#depErr').textContent = 'Select at least one month.'; return; }
  const amount = $('#d_amount').value;
  const d = await api(`/api/tenants/${idx}/deposit`, {method:'POST', body: JSON.stringify({months: _mp.selected, amount})});
  if (d.ok) {
    closeModal();
    toast(d.result.cleared ? `Cleared! ${fmt(d.result.amount)} recorded.` : `${fmt(d.result.amount)} recorded, ${fmt(d.result.new_balance)} left.`);
    state.selectedIdx=idx; state.tab='tenant-detail'; render();
  } else $('#depErr').textContent = d.error || 'Could not record deposit.';
}

function openClearArrears(idx, due) {
  openModal(`
    <h2>Clear Rent Increase Arrears</h2>
    <div class="desc">Outstanding: <b>${fmt(due)}</b>. This is separate from rent and does not affect the due date.</div>
    <label class="field">Method</label>
    <select id="a_method"><option value="Full">Pay in Full</option><option value="Deposit">Partial (Deposit)</option></select>
    <label class="field">Amount (UGX)</label><input id="a_amount" value="${Math.round(due)}">
    <div class="err" id="arrErr"></div>
    <button class="btn btn-primary btn-full" style="margin-top:8px;" onclick="submitArrears(${idx})">Record Payment</button>
  `);
  $('#a_method').addEventListener('change', e => { if (e.target.value==='Full') $('#a_amount').value = Math.round(due); });
}
async function submitArrears(idx) {
  const method = $('#a_method').value, amount = $('#a_amount').value;
  const d = await api(`/api/tenants/${idx}/arrears`, {method:'POST', body: JSON.stringify({method, amount})});
  if (d.ok) { closeModal(); toast('Arrears payment recorded.'); state.selectedIdx=idx; state.tab='tenant-detail'; render(); }
  else $('#arrErr').textContent = d.error || 'Could not record.';
}

// ── UNITS ────────────────────────────────────────────────────────────
async function renderUnits() {
  const d = await api('/api/units');
  $('#headerSub').textContent = `${d.units.length} units`;
  const unitsEmptyMsg = d.no_cache
    ? `<div class="empty"><div class="big">📴</div>You're offline and this hasn't loaded before, so there's nothing cached to show yet.</div>`
    : `<div class="empty"><div class="big">🏢</div>No units yet.</div>`;
  const rows = d.units.map(u => `
    <div class="tenant-row" style="cursor:default;">
      <div class="avatar">🏢</div>
      <div class="meta">
        <div class="name">${escapeHtml(u.name)} ${u.pending_rent_increase?`<span style="font-size:11px;color:var(--warn);font-weight:700;">↑ scheduled</span>`:''}</div>
        <div class="sub">${fmt(u.rent)}/mo · ${u.occupant ? 'Occupied: '+escapeHtml(u.occupant) : 'Vacant'}${u.location?' · '+escapeHtml(u.location):''}</div>
      </div>
      <div style="display:flex;gap:6px;">
        <button class="icon-btn" style="background:var(--teal-soft);color:var(--teal-deep);font-size:13px;width:32px;height:32px;" onclick="openEditUnit('${encodeURIComponent(u.name)}')">✎</button>
        <button class="icon-btn" style="background:var(--teal-soft);color:var(--teal-deep);font-size:13px;width:32px;height:32px;" onclick="openIncreaseRent('${encodeURIComponent(u.name)}', ${u.rent})">↑</button>
      </div>
    </div>`).join('') || unitsEmptyMsg;
  $('#main').innerHTML = `
    <button class="btn btn-primary btn-full" style="margin-bottom:14px;" onclick="openAddUnit()">＋ Add Unit</button>
    <div class="card" style="padding:4px 12px;">${rows}</div>
  `;
}
function openAddUnit() {
  openModal(`
    <h2>Add Unit</h2>
    <label class="field">Unit ID</label><input id="u_name" placeholder="e.g. A1">
    <label class="field">Monthly Rent (UGX)</label><input id="u_rent" inputmode="numeric">
    <label class="field">Location</label><input id="u_location">
    <div class="err" id="unitErr"></div>
    <button class="btn btn-primary btn-full" style="margin-top:8px;" onclick="submitAddUnit()">Save Unit</button>
  `);
}
async function submitAddUnit() {
  const body = {name: $('#u_name').value, rent: $('#u_rent').value, location: $('#u_location').value};
  const d = await api('/api/units', {method:'POST', body: JSON.stringify(body)});
  if (d.ok) { closeModal(); toast('Unit added.'); renderUnits(); } else $('#unitErr').textContent = d.error;
}
async function openEditUnit(nameEnc) {
  const name = decodeURIComponent(nameEnc);
  const d = await api('/api/units');
  const u = d.units.find(x => x.name===name);
  openModal(`
    <h2>Edit Unit — ${escapeHtml(name)}</h2>
    <label class="field">Monthly Rent (UGX)</label><input id="eu_rent" value="${u.rent}">
    <label class="field">Location</label><input id="eu_location" value="${escapeHtml(u.location||'')}">
    <div class="err" id="euErr"></div>
    <button class="btn btn-primary btn-full" style="margin-top:8px;" onclick="submitEditUnit('${nameEnc}')">Save Changes</button>
    <button class="btn btn-danger btn-full" style="margin-top:10px;" onclick="deleteUnit('${nameEnc}')">Delete Unit</button>
  `);
}
async function submitEditUnit(nameEnc) {
  const name = decodeURIComponent(nameEnc);
  const body = {rent: $('#eu_rent').value, location: $('#eu_location').value};
  const d = await api('/api/units/'+encodeURIComponent(name), {method:'PUT', body: JSON.stringify(body)});
  if (d.ok) { closeModal(); toast('Unit updated.'); renderUnits(); } else $('#euErr').textContent = d.error;
}
async function deleteUnit(nameEnc) {
  const name = decodeURIComponent(nameEnc);
  if (!confirm(`Permanently remove unit '${name}'?`)) return;
  const d = await api('/api/units/'+encodeURIComponent(name), {method:'DELETE'});
  if (d.ok) { closeModal(); toast('Unit removed.'); renderUnits(); }
}
function openIncreaseRent(nameEnc, currentRent) {
  const name = decodeURIComponent(nameEnc);
  const months = [];
  const base = new Date(); base.setDate(1);
  for (let i=0;i<24;i++) {
    const m = new Date(base.getFullYear(), base.getMonth()+i, 1);
    const val = m.toISOString().slice(0,10);
    const lbl = m.toLocaleString('default',{month:'short', year:'numeric'});
    months.push(`<option value="${val}" ${i===1?'selected':''}>${lbl}</option>`);
  }
  openModal(`
    <h2>Increase Rent — ${escapeHtml(name)}</h2>
    <div class="desc">Current rent: ${fmt(currentRent)}/month. The occupying tenant's billed rent updates automatically that month.</div>
    <label class="field">New Monthly Rent (UGX)</label><input id="ir_rent">
    <label class="field">Effective From</label><select id="ir_month">${months.join('')}</select>
    <div class="err" id="irErr"></div>
    <button class="btn btn-primary btn-full" style="margin-top:8px;" onclick="submitIncreaseRent('${nameEnc}')">Confirm Increase</button>
  `);
}
async function submitIncreaseRent(nameEnc) {
  const name = decodeURIComponent(nameEnc);
  const body = {new_rent: $('#ir_rent').value, effective_month: $('#ir_month').value};
  const d = await api('/api/units/'+encodeURIComponent(name)+'/increase-rent', {method:'POST', body: JSON.stringify(body)});
  if (d.ok) { closeModal(); toast('Rent increase scheduled.'); renderUnits(); } else $('#irErr').textContent = d.error;
}

// ── ALERTS ───────────────────────────────────────────────────────────
async function renderAlerts() {
  const d = await api('/api/alerts');
  $('#headerSub').textContent = `${d.alerts.length} tenant(s) to watch`;
  const alertsEmptyMsg = d.no_cache
    ? `<div class="empty"><div class="big">📴</div>You're offline and this hasn't loaded before, so there's nothing cached to show yet.</div>`
    : `<div class="empty"><div class="big">✅</div>No alerts. Everyone's paid up.</div>`;
  const rows = d.alerts.map(tenantRowHtml).join('') || alertsEmptyMsg;
  $('#main').innerHTML = `<div class="section-title">Overdue &amp; Upcoming</div><div class="card" style="padding:4px 12px;">${rows}</div>`;
}

// ── SETTINGS ─────────────────────────────────────────────────────────
function renderMoreMenu() {
  $('#headerSub').textContent = 'More';
  $('#main').innerHTML = `
    <div class="card" style="padding:0;overflow:hidden;">
      <div class="menu-row" onclick="switchTab('history')">
        <span class="menu-icon">📜</span>
        <div class="menu-text">
          <div class="menu-title">Transaction History</div>
          <div class="menu-sub">Every tenant's payment &amp; deposit ledger</div>
        </div>
        <span class="menu-chevron">›</span>
      </div>
      <div class="menu-row" onclick="switchTab('settings')">
        <span class="menu-icon">⚙️</span>
        <div class="menu-text">
          <div class="menu-title">Settings</div>
          <div class="menu-sub">App, security, reports &amp; data</div>
        </div>
        <span class="menu-chevron">›</span>
      </div>
    </div>
  `;
}

async function renderHistory() {
  const q = state.historyQ || '';
  $('#headerSub').textContent = 'Transaction History';
  const [d, dash] = await Promise.all([
    api('/api/history' + (q ? ('?q=' + encodeURIComponent(q)) : '')),
    api('/api/dashboard'),
  ]);
  const tenants = d.tenants || [];

  const incomeCard = `<div class="card" style="background:var(--teal-soft2);border-color:var(--teal-soft);">
    <div class="section-title" style="margin-top:0;">Monthly Income — ${escapeHtml(dash.month_name || '')}</div>
    <div style="font-family:'Space Grotesk';font-size:24px;font-weight:700;color:var(--teal-deep);">${fmt(dash.month_income || 0)}</div>
    <div style="display:flex;gap:16px;margin-top:8px;font-size:12px;">
      <div><b style="color:var(--good);">${fmt(dash.full_payment_total || 0)}</b><div style="color:var(--muted);">Full</div></div>
      <div><b style="color:var(--teal-deep);">${fmt(dash.deposit_total || 0)}</b><div style="color:var(--muted);">Deposits</div></div>
      <div><b style="color:var(--danger);">${fmt(dash.cancelled_total || 0)}</b><div style="color:var(--muted);">Cancelled</div></div>
    </div>
  </div>`;

  const monthlyCard = `<div class="section-title">Monthly Transactions</div>
  <div class="card">
    <div class="sub" style="color:var(--muted);font-size:12.5px;margin-bottom:10px;">
      Generate a summary of payments and deposits for one specific month.
    </div>
    <div style="display:flex;gap:8px;">
      <div class="select-native-wrap"><select id="mrMonth"></select></div>
      <div class="select-native-wrap"><select id="mrYear"></select></div>
    </div>
    <button class="btn btn-primary btn-full" style="margin-top:10px;" onclick="generateMonthlyReport()">Generate</button>
    <div id="mrResults" style="margin-top:12px;"></div>
  </div>`;

  const rowsHtml = (t) => {
    if (!t.transactions.length) {
      return `<div class="hist-empty">No transactions recorded yet.</div>`;
    }
    return `<div class="hist-table">
      <div class="hist-hdr">
        <div>Date</div><div>Type</div><div>Amount</div><div>Period</div>
      </div>
      ${t.transactions.map(rec => {
        const cancelled = rec.cancelled;
        const typeStr = (rec.kind === 'deposit' ? 'Deposit' : 'Full') + (cancelled ? '  ✕' : '');
        const rowCls = cancelled ? 'hist-row hist-cancelled' : 'hist-row';
        return `<div class="${rowCls}">
          <div>${escapeHtml(rec.date)}</div>
          <div>${typeStr}</div>
          <div>${fmt(rec.amount)}</div>
          <div class="hist-period">${escapeHtml(abbrevPeriod(rec.from, rec.to))}</div>
        </div>`;
      }).join('')}
    </div>`;
  };

  const cardsHtml = tenants.length
    ? tenants.map(t => `
      <div class="card">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;">
          <div>
            <div style="font-weight:700;font-size:15px;">${escapeHtml(t.name)}</div>
            <div class="sub" style="color:var(--muted);font-size:12.5px;">🏠 ${escapeHtml(t.unit)}</div>
          </div>
          <button class="btn btn-ghost" style="padding:6px 10px;font-size:12px;" onclick="openTenant(${t.index})">Profile</button>
        </div>
        <div style="margin-top:10px;font-size:11px;color:var(--muted);text-transform:uppercase;font-weight:700;">Entry Date</div>
        <div style="font-size:14px;font-weight:700;margin-top:2px;">${escapeHtml(t.entry_date)}</div>
        <div style="margin-top:12px;font-size:11px;color:var(--muted);text-transform:uppercase;font-weight:700;">Transaction History</div>
        <div style="margin-top:8px;">${rowsHtml(t)}</div>
      </div>`).join('')
    : `<div class="empty">No tenants match your search.</div>`;

  $('#main').innerHTML = `
    <button class="btn btn-ghost" style="margin-bottom:12px;" onclick="switchTab('more');">← Back</button>
    <div class="searchbar"><span>🔎</span><input id="historySearch" placeholder="Search name, unit or phone…" value="${escapeHtml(q)}"></div>
    ${incomeCard}
    ${monthlyCard}
    ${cardsHtml}
  `;
  $('#historySearch').addEventListener('input', debounce(e => { state.historyQ = e.target.value; renderHistory(); }, 300));
  populateMonthlyPickers();
}

async function renderSettings() {
  $('#headerSub').textContent = 'Settings';
  $('#main').innerHTML = `
    <button class="btn btn-ghost" style="margin-bottom:12px;" onclick="switchTab('more');">← Back</button>
    <div class="section-title">Get the App</div>
    <div class="card" id="installCard">
      <div class="sub" style="color:var(--muted);font-size:12.5px;margin-bottom:10px;">
        Save this to your phone's home screen or browser for one-tap access — no need to retype the address each time.
      </div>
      <button class="btn btn-primary btn-full" id="installBtn" onclick="triggerInstall()" style="display:none;">📲 Install App</button>
      <div class="sub" id="installHint" style="color:var(--muted);font-size:11.5px;margin-top:8px;"></div>
    </div>
    <div class="section-title">Reports</div>
    <div class="card">
      <a class="linklike" href="api/export/excel">⬇ Download Excel (.xlsx)</a><br><br>
      <a class="linklike" href="api/export/pdf">⬇ Download PDF Report</a>
    </div>
    <div class="section-title">Danger Zone</div>
    <div class="card">
      <div class="sub" style="color:var(--muted);font-size:12.5px;margin-bottom:10px;">
        Permanently erases every tenant, unit, and transaction shared with the desktop app. A backup of the current data is kept automatically before wiping, and the PC's admin PIN is required.
      </div>
      <button class="btn btn-danger btn-full" onclick="openResetData()">🗑 Reset All Data</button>
    </div>
    <div class="section-title">About</div>
    <div class="card sub" style="color:var(--muted);font-size:12.5px;">
      Tenant Monitoring &amp; Management — Web Edition.<br>Shares data with the desktop app on this PC.<br><br>
      Device naming, connected-device management, and the app lock/PIN are all controlled from the desktop app now, not from here.
    </div>
  `;
  updateInstallUI();
}

// ── DEVICE NAMING (prompted once, right after a QR pairing scan) ──────
// PIN/lock and the connected-devices roster are admin-only now and are
// managed from the desktop app's Settings -> Connect Phone panel instead
// of from here -- a phone can still name ITSELF (device_id is already
// tied to it, so this is just a courtesy so the admin can tell devices
// apart in that PC-side panel), but nothing else about security or the
// device list is editable from the web app.
async function maybePromptDeviceName() {
  const justPaired = new URLSearchParams(location.search).has('pt');
  if (!justPaired) return;
  try {
    const d = await api('/api/devices');
    const devices = (d && d.devices) || [];
    const mine = devices.find(dev => dev.device_id === DEVICE_ID);
    if (mine && mine.label) return;  // already named (e.g. re-scanned an old QR)
  } catch (e) { return; }
  openModal(`
    <h2>Name This Device</h2>
    <div class="desc">Give this phone/browser a name so the admin can recognize it in the connected-devices list on the PC.</div>
    <label class="field">Device Name</label>
    <input id="dev_my_label" placeholder="e.g. Mary's iPhone">
    <div class="err" id="devLabelErr"></div>
    <button class="btn btn-primary btn-full" style="margin-top:8px;" onclick="submitDeviceLabel()">Save Name</button>
  `, {dismissible: false});
}
async function submitDeviceLabel() {
  const label = $('#dev_my_label').value.trim();
  if (!label) { $('#devLabelErr').textContent = 'Please enter a name.'; return; }
  const d = await api('/api/devices/label', {method:'POST', body: JSON.stringify({label})});
  if (d.ok) { closeModal(); toast('Device name saved.'); }
  else $('#devLabelErr').textContent = d.error || 'Could not save name.';
}

// ── DANGER ZONE: reset data ──────────────────────────────────────────
function openResetData() {
  openModal(`
    <h2>Reset All Data</h2>
    <div class="desc">This permanently deletes every tenant, unit, and transaction on both this app and the desktop app. A backup of the current data file is saved automatically first. This cannot be undone from here.</div>
    <label class="field">Admin PIN (set on the PC)</label><input id="reset_admin_pin" type="password" inputmode="numeric">
    <div class="err" id="resetErr"></div>
    <button class="btn btn-danger btn-full" style="margin-top:8px;" onclick="submitResetData()">Yes, Reset Everything</button>
  `);
}
async function submitResetData() {
  const adminPin = $('#reset_admin_pin').value.trim();
  if (!adminPin) { $('#resetErr').textContent = 'Enter the admin PIN set on the PC.'; return; }
  if (!confirm('Are you absolutely sure? All tenants and units will be erased.')) return;
  const d = await api('/api/settings/reset', {method:'POST', body: JSON.stringify({admin_pin: adminPin})});
  if (d.ok) { closeModal(); toast('All data reset.'); cacheDeletePrefix('/api/'); switchTab('dashboard'); }
  else $('#resetErr').textContent = d.error || 'Could not reset data.';
}

const MONTH_NAMES = ['January','February','March','April','May','June','July',
                      'August','September','October','November','December'];

function populateMonthlyPickers() {
  const now = new Date();
  const monthSel = $('#mrMonth'), yearSel = $('#mrYear');
  if (!monthSel || !yearSel) return;
  monthSel.innerHTML = MONTH_NAMES.map((m, i) =>
    `<option value="${i+1}" ${i === now.getMonth() ? 'selected' : ''}>${m}</option>`).join('');
  const startYear = now.getFullYear() - 3;
  let yearsHtml = '';
  for (let y = startYear; y <= now.getFullYear() + 1; y++) {
    yearsHtml += `<option value="${y}" ${y === now.getFullYear() ? 'selected' : ''}>${y}</option>`;
  }
  yearSel.innerHTML = yearsHtml;
}

async function generateMonthlyReport() {
  const month = $('#mrMonth').value, year = $('#mrYear').value;
  const box = $('#mrResults');
  box.innerHTML = `<div class="sub" style="color:var(--muted);">Generating…</div>`;
  try {
    const res = await fetchTimeout(`/api/monthly-report?year=${year}&month=${month}`, {}, 6000);
    const report = await res.json();
    if (report.error) { box.innerHTML = `<div class="sub" style="color:var(--danger);">${escapeHtml(report.error)}</div>`; return; }
    const rowsHtml = report.tenant_rows.length
      ? report.tenant_rows.map(r => `
        <div class="hist-row">
          <div>${escapeHtml(r.name)} <span style="color:var(--muted);">(${escapeHtml(r.unit)})</span></div>
          <div style="text-align:right;font-weight:700;">${fmt(r.pay_active + r.dep_active)}</div>
        </div>`).join('')
      : `<div class="hist-empty">No transactions in ${escapeHtml(report.month_label)}.</div>`;
    box.innerHTML = `
      <div style="font-family:'Space Grotesk';font-size:20px;font-weight:700;color:var(--teal-deep);margin-top:6px;">
        ${fmt(report.grand_combined)}
      </div>
      <div style="display:flex;gap:16px;margin:6px 0 12px;font-size:12px;">
        <div><b style="color:var(--good);">${fmt(report.grand_pay)}</b><div style="color:var(--muted);">Full</div></div>
        <div><b style="color:var(--teal-deep);">${fmt(report.grand_dep)}</b><div style="color:var(--muted);">Deposits</div></div>
        <div><b style="color:var(--danger);">${fmt(report.grand_cancelled)}</b><div style="color:var(--muted);">Cancelled</div></div>
      </div>
      ${rowsHtml}
      <a class="linklike" style="display:block;margin-top:12px;" href="api/export/monthly-excel?year=${year}&month=${month}">⬇ Download Monthly Excel (.xlsx)</a>
    `;
  } catch (e) {
    box.innerHTML = `<div class="sub" style="color:var(--danger);">Couldn't generate the report — check the connection and try again.</div>`;
  }
}

// ── pull-to-refresh ──────────────────────────────────────────────────
// Only engages when the page is scrolled all the way to the top (so it
// never fights with normal scrolling further down), and re-runs whatever
// tab is currently showing by re-dispatching switchTab, which re-fetches
// from the server (bypassing nothing — /api/* is never cached by the
// service worker, see sw.js above) and re-renders with fresh data.
(function setupPullToRefresh() {
  const main = $('#main');
  const indicator = $('#ptrIndicator');
  if (!main || !indicator) return;
  const THRESHOLD = 70;
  let startY = null, pulling = false, refreshing = false;

  main.addEventListener('touchstart', (e) => {
    if (refreshing || main.scrollTop > 0) { startY = null; return; }
    startY = e.touches[0].clientY;
    pulling = true;
  }, { passive: true });

  main.addEventListener('touchmove', (e) => {
    if (!pulling || startY === null || refreshing) return;
    const dy = e.touches[0].clientY - startY;
    if (dy <= 0) return;
    indicator.classList.add('visible');
    if (dy > THRESHOLD) {
      indicator.classList.add('ready');
      indicator.textContent = '↑ Release to refresh';
    } else {
      indicator.classList.remove('ready');
      indicator.textContent = '↓ Pull to refresh';
    }
  }, { passive: true });

  main.addEventListener('touchend', async (e) => {
    if (!pulling || startY === null || refreshing) { pulling = false; return; }
    const ready = indicator.classList.contains('ready');
    pulling = false;
    startY = null;
    if (!ready) {
      indicator.classList.remove('visible', 'ready');
      return;
    }
    refreshing = true;
    indicator.textContent = '⟳ Refreshing…';
    indicator.classList.add('spinning');
    await pingServer();
    switchTab(state.tab);
    setTimeout(() => {
      indicator.classList.remove('visible', 'ready', 'spinning');
      refreshing = false;
    }, 400);
  });
})();

// ── boot ─────────────────────────────────────────────────────────────
async function boot() {
  booted = true;
  loadCloudConfig();
  switchTab('dashboard');
  maybePromptDeviceName();
}
async function init() {
  let ls;
  try {
    const res = await fetchTimeout('/api/lock-status', {headers: {'X-Device-Id': DEVICE_ID}}, 3500);
    ls = await res.json();
    adoptCanonicalDeviceId(ls && ls.canonical_device_id);
    if (ls && (ls.kicked || ls.disconnecting || ls.pending_approval || ls.device_limit_reached)) {
      enterBlockedState(ls);
      return;
    }
    cacheSet('/api/lock-status', ls);
    setOnline(true);
  } catch (err) {
    setOnline(false);
    ls = cacheGet('/api/lock-status') || { pin_set: false, unlocked: true };
  }
  if (ls.pin_set && !ls.unlocked) { showLock(); }
  else { hideLock(); boot(); }
  updateSyncBadge();
}
init();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    no_browser = os.environ.get("RM_NO_BROWSER") == "1"
    print(f"\n  {APP_NAME} -- web edition (single file)")
    print(f"  Data file: {DATA_FILE}")
    print(f"  Local:  http://127.0.0.1:{port}")
    if CLOUD_MODE:
        print("  Running in CLOUD_MODE -- reachable at this service's public URL; "
              "phones pair to it directly via the QR code the desktop app shows "
              "under Settings -> Connect Phone.\n")
    else:
        print("  Not running in CLOUD_MODE (no DATABASE_URL) -- this is a local "
              "test run only. Deploy with CLOUD_MODE=1 and a DATABASE_URL for "
              "phones to be able to reach it.\n")

    if not no_browser:
        # Standalone run (person double-clicked/ran this file themselves) —
        # open the app on this PC as a convenience.
        def _open_local_page():
            try:
                webbrowser.open(f"http://127.0.0.1:{port}")
            except Exception:
                pass
        threading.Timer(1.0, _open_local_page).start()
    # else: launched from the desktop app's Settings → Connect Phone, which
    # already shows its own QR code in-window — nothing should open on the PC.

    app.run(host="0.0.0.0", port=port, debug=False)
