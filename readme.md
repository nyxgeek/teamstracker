# TeamsTracker

tool from dc31 talk - Track the Planet. See slides here: https://github.com/nyxgeek/track_the_planet

## Overview

```
usage: teamstracker.py [-h] [-c CSV] [-u UUID] [-U UUIDFILE] [-v] [-T THREADS]

optional arguments:
  -h, --help            show this help message and exit
  -c CSV, --csv CSV     csv export of email addresses from TeamFiltration
  -u UUID, --uuid UUID  Azure UUID to target
  -U UUIDFILE, --uuidfile UUIDFILE
                        file containing Azure UUIDs of users to target
  -v, --verbose         enable verbose output
  -T THREADS, --threads THREADS
                        total number of threads (defaut: 50)
```



## to disable external access:

https://learn.microsoft.com/en-us/microsoftteams/trusted-organizations-external-meetings-chat?tabs=organization-settings
