# Olm

Emacs interface for MS Outlook Mail


## About

Olm is an Emacs interface for MS Outlook mails. It consists of two
parts, an Emacs interface written in Emacs Lisp and glue scripts
written in Ruby.  It is currently designed against MS Outlook 2007.


## Screenshots

![olm-message.png](https://qiita-image-store.s3.amazonaws.com/0/12281/55b58dc1-ccad-e6e4-3682-c9b416a0965c.png)

![olm-draft.png](https://qiita-image-store.s3.amazonaws.com/0/12281/cecea40d-e898-e8fd-2bd5-3fa9a3afc57e.png)


## Requirements

+ Emacs 24.3 or later
+ [RubyInstaller](http://rubyinstaller.org/) 2.0.0-p598
+ [olm](http://rubygems.org/gems/olm) gem
+ [dash.el](https://github.com/magnars/dash.el)


## Installation


**olm gem**

Install the olm gem using `gem install` command.

    $ gem install olm


**olm.el**

This package has not been part of MELPA or Marmalade. You have to install `olm.el` manually.

+ `git clone https://github.com/tnoda/olm.git`
+ `M-x package-install-file RET`
+ Select the `olm.el`'s path that you git-cloned


# Usage

Once installed, `M-x olm` to enter `olm-summary-mode`.

`olm-summary-mode` displays the list of emails in your inbox,
where you can use olm commands by pressing the following keys.

- `.` (`olm-summary-open-message`):
  - Open the message at point.

- `d` (`olm-summary-delete`):
  - Delete the message at point.

- `w` (`olm-summary-write`):
  - Create a new email draft buffer. You can save and send the buffer content
    by `C-c C-s` and `C-c C-c`, respectively.

- `g` (`olm-summary-goto-folder`):
  - Go to a folder you specify.
- `o` (`olm-summary-refile`):
  - Put the refile mark ('o') on a message at point.
- `x` (`olm-summary-exec`):
  - Process marked messages.


# See Also

Olm does not have its own address book. Instead, it is designed to work well with [helm-ad](https://github.com/tnoda/helm-ad), helm source for Active Directory.


# Disclaimer

This project has been abandoned since I stopped using MS Outlook. It
is not updated and maintained any more. Please feel free to fork it.
