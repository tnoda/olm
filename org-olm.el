;;; org-olm.el --- Org links to Outlook Mail messages

;; Copyright (C) 2014  Takahiro Noda

;; Author: Takahiro Noda <takahiro.noda+github@gmail.com>
;; Keywords: mail

;; This program is free software; you can redistribute it and/or modify
;; it under the terms of the GNU General Public License as published by
;; the Free Software Foundation, either version 3 of the License, or
;; (at your option) any later version.

;; This program is distributed in the hope that it will be useful,
;; but WITHOUT ANY WARRANTY; without even the implied warranty of
;; MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
;; GNU General Public License for more details.

;; You should have received a copy of the GNU General Public License
;; along with this program.  If not, see <http://www.gnu.org/licenses/>.

;;; Commentary:

;; This file implements links to Outlook Mail messages from within Org mode.
;; 

;;; Code:

(require 'org)
(require 'olm)

;;; Install the link type
(org-add-link-type "olm" 'org-olm-open)
(add-hook 'org-store-link-functions 'org-olm-store-link)

;;; Implementation
(defun org-olm-open (path)
  "Follow an Outlook Mail message link specified by PATH."
  (error "org-olm-open has not been implemented yet"))

(defun org-olm-store-link ()
  "Store a link to an Outlook Mail message."
  (or (org-olm-store-link-summary)
      (org-olm-store-link-message)))

(defun org-olm-store-link-summary ()
  (when (eq major-mode 'olm-summary-mode)
    (org-store-link-props
     :type "olm"
     :link (concat "olm:" (olm-mail-item-entry-id-at))
     :description (org-olm-store-link-description))))

(defun org-olm-store-link-message ()
  (when (eq major-mode 'olm-message-mode)
    (error "org-olm-store-link-message has not been implemented yet")))

(defun org-olm-store-link-description ()
  (with-current-buffer (get-buffer "*olm-message*")
    (save-excursion
      (save-restriction
        (let ((from (progn
                      (goto-char (point-min))
                      (re-search-forward "^From: *\\(.+\\)" nil t)
                      (match-string-no-properties 1)))
              (subj (progn
                      (goto-char (point-min))
                      (re-search-forward "^Subject: *\\(.+\\)" nil t)
                      (match-string-no-properties 1))))
          (concat "olm:" from "/" subj))))))

(provide 'org-olm)
;;; org-olm.el ends here
