# cronie

Simple tool for running cronjobs without crond.

## Table of contents

<!-- vim-markdown-toc GFM -->

* [Install](#install)
* [Usage](#usage)

<!-- vim-markdown-toc -->

## Install

```bash
npm global install cronie
```

## Usage

Run `curl google.com` when time matches "\* \* \* \* \*"

```bash
cronie run "* * * * *" curl google.com
```

Run `ping google.com` and restart it when time matches "\* \* \* \* \*"

```bash
cronie restart "* * * * *" ping google.com
```

