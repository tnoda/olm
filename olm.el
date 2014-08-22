;;; olm.el --- Outlook Mail in Emacs

;; Copyright (C) 2014  Takahiro Noda

;; Author: Takahiro Noda <takahiro.noda+github@gmail.com>
;; Keywords: mail
;; Package-Requires: ((dash "2.8.0"))

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

;; Outlook Mail in Emacs

;;; Code:

(require 'dash)

(defvar olm-folder-id nil)
(defvar olm-folder-name nil)
(defvar olm-default-bcc nil)
(defvar olm-attachment-path nil)
(defvar olm-ruby-executable "ruby")
(defvar olm-deleted-items-folder-id nil)

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; olm-folder-alist
;;; 
(defvar olm-folder-alist nil)

(defun olm-folder-names ()
  (--map (car it) olm-folder-alist))

(unless olm-folder-id
  (setq olm-folder-id (cadr olm-folder-alist)))
(unless olm-folder-name
  (setq olm-folder-name (caar olm-folder-alist)))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;;  olm command
;;; 
(defun olm ()
  (interactive)
  (olm-init)
  (olm-scan))

(defun olm-init ()
  (unless olm-folder-id
    (setq olm-folder-id (cadr olm-folder-alist)
          olm-folder-name (caar olm-folder-alist)))
  (unless olm-folder-id
    (setq olm-folder-id (olm-default-folder-id)
          olm-folder-name "inbox"))
  (unless olm-deleted-items-folder-id
    (setq olm-deleted-items-folder-id (olm-deleted-items-folder-id))))

(defun olm-scan ()
  (interactive)
  (let* ((ret (olm-ls))
         (lbuf (olm-buf-ls))
         (n (with-current-buffer lbuf
              (goto-char (point-min))
              (let ((point (point)))
                (forward-word)
                (string-to-int (buffer-substring point (point))))))
         (sbuf (let ((buf (olm-buf-summary)))
                 (with-current-buffer buf
                   (setq-local buffer-read-only nil)
                   (erase-buffer))
                 buf))
         (ibuf (olm-buf-entry-ids t)))
    (with-current-buffer lbuf
      (--dotimes n
        (progn
          (forward-line)
          (let* ((p0 (point))
                 (p1 (progn
                       (forward-word)
                       (point)))
                 (entry-id (buffer-substring p0 p1))
                 (item-line (buffer-substring p1 (line-end-position))))
            (with-current-buffer sbuf
              (insert item-line "\n"))
            (with-current-buffer ibuf
              (insert entry-id "\n"))))))
    (with-current-buffer sbuf
      (olm-summary-mode))
    (switch-to-buffer sbuf)))

;;; A helper function invoked by olm-scan
(defun olm-ls ()
  (with-current-buffer (olm-buf-ls)
    (setq-local buffer-read-only nil)
    (erase-buffer)
    (let ((command (if olm-folder-id
                       (format "Olm.ls %S" olm-folder-id)
                     "Olm.ls")))
      (olm-do-command command t))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; general helper functions
;;; 
(defun olm-do-command (command &optional buf)
  (let ((buf (or buf (get-buffer-create "*Messages*"))))
    (call-process olm-ruby-executable nil buf nil
                  "-r" "rubygems" "-r" "olm" "-e" command)))

(defun olm-do-command-buf (command)
  (call-process-region (point-min) (point-max)
                       olm-ruby-executable
                       nil (get-buffer-create "*Messages*") nil
                       "-r" "rubygems" "-r" "olm" "-e" command))

(defun olm-sync ()
  (interactive)
  (with-current-buffer (get-buffer-create "*olm-sync*")
    (message "Olm: synchronizing all objects ...")
    (olm-do-command "Olm.send_and_receive" t)
    (message "Olm: synchronizing all objects ...")
    (sit-for 1)
    (let ((w 5))
      (--dotimes w
        (progn
          (message (number-to-string (- w it)))
          (sit-for 1))))
    (message "done.")))

(defun olm-hide-entry-id-line ()
  (interactive)
  (let ((inhibit-read-only t))
    (put-text-property (progn
                         (goto-char (point-min))
                         (point))
                       (line-end-position)
                       'read-only
                       nil))
  (save-excursion    
    (narrow-to-region (progn
                        (goto-line 2)
                        (point))
                      (point-max))))

(defun olm-default-folder-id ()
  (with-current-buffer (get-buffer-create "*olm-default-folder-id*")
    (erase-buffer)
    (olm-do-command "Olm.default_folder_id" (current-buffer))
    (buffer-substring-no-properties (point-min) (1- (point-max)))))

(defun olm-deleted-items-folder-id ()
  (with-current-buffer (get-buffer-create "*olm-deleted-items-folder-id*")
    (erase-buffer)
    (olm-do-command "Olm.deleted_items_folder_id" (current-buffer))
    (buffer-substring-no-properties (point-min) (1- (point-max)))))

(defun olm-pick-a-folder ()
  (assoc-default (completing-read "Refile to: "
                                  (olm-folder-names)
                                  nil
                                  t)
                 olm-folder-alist))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; buffers
;;; 
(defun olm-buf-ls ()
  (get-buffer-create "*olm-ls*"))

(defun olm-buf-summary ()
  (let ((buf (get-buffer-create "*olm-summary*")))
    (with-current-buffer buf
      (setq-local inherit-process-coding-system t))
    buf))

(defun olm-buf-entry-ids (&optional erase-buffer)
  (let ((buf(get-buffer-create "*olm-entry-ids*")))
    (when erase-buffer
      (with-current-buffer buf
        (setq-local buffer-read-only nil)
        (erase-buffer)))
    buf))

(defun olm-buf-message ()
  (let ((buf (get-buffer-create "*olm-message*")))
    (with-current-buffer buf
      (setq-local buffer-read-only nil)
      (erase-buffer))
    buf))

(defun olm-buf-draft ()
  (let ((buf (get-buffer-create "*olm-draft*")))
    (with-current-buffer buf
      (setq-local buffer-read-only nil)
      (erase-buffer)
      (insert "\n")
      (insert "To: \n")
      (insert "Cc: \n")
      (when olm-default-bcc
        (insert "Bcc: " olm-default-bcc "\n"))
      (insert "Subject: \n")
      (insert "---- \n"))
    buf))

(defun olm-buf-draft-reply-all (entry-id)
  (let ((buf (generate-new-buffer "*olm-draft-reply-all*")))
    (with-current-buffer buf
      (olm-do-command (format "Olm.create_reply_all_message %S"
                              entry-id)
                      t)
      (goto-char (point-min))
      (re-search-forward "^From: ")
      (insert "***Reply All***")
      (when olm-default-bcc
        (goto-char (point-min))
        (re-search-forward "^---- ")
        (beginning-of-line)
        (insert "Bcc: " olm-default-bcc "\n"))
      (olm-draft-reply-all-mode)
      (let ((inhibit-read-only t))
        (put-text-property (progn
                             (goto-char (point-min))
                             (point))
                           (progn
                             (re-search-forward "^---- ")
                             (point))
                           'read-only
                           t))
      (olm-hide-entry-id-line)
      (forward-line))
    buf))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; olm-message
;;; 
(defvar olm-message-mode-map nil)
(defvar olm-message-mode-hook nil)

(unless olm-message-mode-map
  (setq olm-message-mode-map (make-sparse-keymap))
  (define-key olm-message-mode-map "\C-c\C-a" 'olm-message-save-attachments)
  (define-key olm-message-mode-map "\C-c\C-r" 'olm-message-reply-all))

(defun olm-message-mode ()
  (interactive)
  (use-local-map olm-message-mode-map)
  (setq major-mode 'olm-message-mode)
  (setq mode-name (format "Olm Message" olm-folder-name))
  (font-lock-mode 1)
  (setq-local buffer-read-only t)
  (setq-local line-move-ignore-invisible t)
  (run-hooks 'olm-message-mode-hook))

(defun olm-message-mode-keyword ()
  (font-lock-add-keywords
   nil
   '(("^From:" . font-lock-keyword-face)
     ("^To:" . font-lock-keyword-face)
     ("^Cc:" . font-lock-keyword-face)
     ("^Subject:" . font-lock-keyword-face)
     ("^ReceivedAt:" . font-lock-keyword-face)
     ("^Attachments:" . font-lock-keyword-face)
     ("^> .*$" . font-lock-comment-face)
     ("^From: \\(.*\\)$" 1 font-lock-warning-face)
     ("^To: \\(.*\\)$" 1 font-lock-negation-char-face)
     ("^Cc: \\(.*\\)$" 1 font-lock-constant-face)
     ("^Subject: \\(.*\\)$" 1 font-lock-variable-name-face)
     ("^ReceivedAt: \\(.*\\)$" 1 font-lock-type-face))))

(add-hook 'olm-message-mode-hook 'olm-message-mode-keyword)

(defun olm-message-save-attachments ()
  (interactive)
  (let ((entry-id (olm-message-entry-id)))
    (message (format "Olm: saving attachments into %S ..."
                     olm-attachment-path))
    (olm-do-command (format "Olm.save_attachments(%S, %S)"
                            entry-id
                            olm-attachment-path))))

(defun olm-message-reply-all ()
  (interactive)
  (-> (olm-message-entry-id)
    olm-buf-draft-reply-all
    switch-to-buffer))


;;; A helper function for olm-message-mode
(defun olm-message-entry-id ()
  (interactive)
  (save-excursion
    (save-restriction
      (widen)
      (goto-line 1)
      (buffer-substring-no-properties (line-beginning-position)
                                      (line-end-position)))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; olm-summary
;;;
(defvar olm-summary-mode-map nil)
(defvar olm-summary-mode-hook nil)
(defvar olm-summary-scroll-lines 12)

(unless olm-summary-mode-map
  (setq olm-summary-mode-map (make-sparse-keymap))
  (define-key olm-summary-mode-map "s" 'olm-scan)
  (define-key olm-summary-mode-map "i" 'olm-summary-inc)
  (define-key olm-summary-mode-map "q" 'olm-summary-quit)
  (define-key olm-summary-mode-map "." 'olm-summary-open-message)
  (define-key olm-summary-mode-map " " 'olm-summary-scroll-message-forward)
  (define-key olm-summary-mode-map "\r" 'olm-summary-scroll-message-forward-line)
  (define-key olm-summary-mode-map [backspace] 'olm-summary-scroll-message-backward)
  (define-key olm-summary-mode-map "p" 'olm-summary-display-up)
  (define-key olm-summary-mode-map "n" 'olm-summary-display-down)
  (define-key olm-summary-mode-map "!" 'olm-summary-toggle-flag)
  (define-key olm-summary-mode-map "w" 'olm-summary-write)
  (define-key olm-summary-mode-map "A" 'olm-summary-reply-all)
  (define-key olm-summary-mode-map "g" 'olm-summary-goto-folder)
  (define-key olm-summary-mode-map "o" 'olm-summary-refile)
  (define-key olm-summary-mode-map "d" 'olm-summary-delete)
  (define-key olm-summary-mode-map "x" 'olm-summary-exec)
  (define-key olm-summary-mode-map "*" 'olm-summary-review)
  (define-key olm-summary-mode-map "mo" 'olm-summary-mark-refile))

(defun olm-summary-inc ()
  (interactive)
  (olm-sync)
  (olm-scan))

(defun olm-summary-quit ()
  (interactive)
  (delete-other-windows-vertically)
  (quit-window))

(defun olm-summary-mode ()
  (interactive)
  (use-local-map olm-summary-mode-map)
  (setq major-mode 'olm-summary-mode)
  (setq mode-name (format "Olm Summary [%s]" olm-folder-name))
  (font-lock-mode 1)
  (setq-local buffer-read-only t)
  (setq-local truncate-lines t)
  (setq-local line-move-ignore-invisible t)
  (run-hooks 'olm-summary-mode-hook))

(defun olm-summary-mode-keyword ()
  (font-lock-add-keywords
   nil
   '(("^o.*$" . font-lock-builtin-face)
     ("^D.*$" . font-lock-warning-face)
     ("^\\*.*$" . font-lock-comment-face))))

(add-hook 'olm-summary-mode-hook 'olm-summary-mode-keyword)

(defun olm-summary-open-message ()
  (interactive)
  (delete-other-windows-vertically)
  (let* ((ln0 (line-number-at-pos))
         (msg-window (split-window-below 12))
         (entry-id (progn
                     (goto-line ln0)
                     (olm-summary-message-entry-id))))
    (olm-open-message entry-id msg-window))
  (olm-summary-mark-message-as-read))

(defun olm-summary-mark-message-as-read ()
  (interactive)
  (olm-do-command (format "Olm.mark_as_read %S" (olm-summary-message-entry-id))))

(defun olm-summary-scroll-message-forward ()
  (interactive)
  (recenter)
  (scroll-other-window olm-summary-scroll-lines))

(defun olm-summary-scroll-message-forward-line ()
  (interactive)
  (recenter)
  (scroll-other-window 1))

(defun olm-summary-scroll-message-backward ()
  (interactive)
  (recenter)
  (scroll-other-window (- olm-summary-scroll-lines)))

(defun olm-summary-display-up ()
  (interactive)
  (forward-line -1)
  (recenter)
  (olm-summary-open-message))

(defun olm-summary-display-down ()
  (interactive)
  (forward-line 1)
  (recenter)
  (olm-summary-open-message))

(defun olm-summary-toggle-flag ()
  (interactive)
  (let ((n (line-number-at-pos)))
    (olm-do-command (format "Olm.toggle_task_flag %S"
                            (olm-summary-message-entry-id)))
    (olm-scan)
    (goto-line n)))

(defun olm-summary-write ()
  (interactive)
  (delete-other-windows-vertically)
  (let ((buf (olm-buf-draft)))
    (with-current-buffer buf
      (olm-hide-entry-id-line)
      (olm-draft-mode))
    (switch-to-buffer buf)))

(defun olm-summary-reply-all ()
  (interactive)
  (let ((buf (olm-buf-draft-reply-all (olm-summary-message-entry-id))))
    (delete-other-windows-vertically)
    (switch-to-buffer buf)))

(defun olm-summary-goto-folder ()
  (interactive)
  (setq olm-folder-name (completing-read "Exchange folder: "
                                         (olm-folder-names)
                                         nil
                                         t))
  (setq olm-folder-id (assoc-default olm-folder-name olm-folder-alist))
  (olm-scan))

(defun olm-summary-refile (&optional to mark)
  (interactive)
  (let ((from (olm-summary-message-entry-id))
        (to (or to (olm-pick-a-folder)))
        (mark (or mark "o"))
        (n (line-number-at-pos)))
    ;; insert the destination entry id into the org-ids buffer
    (with-current-buffer (olm-buf-entry-ids)
      (goto-line n)
      (re-search-forward "^[0-9A-Z]\\{140\\}")
      (delete-region (point) (point-at-eol))
      (insert " " to))
    (olm-summary-insert-mark mark)))

(defun olm-summary-delete ()
  (interactive)
  (olm-summary-refile olm-deleted-items-folder-id "D"))

(defun olm-summary-exec ()
  "Process marked messages."
  (interactive)
  (with-current-buffer (olm-buf-entry-ids)
    (olm-do-command-buf "Olm.execute_refile"))
  (olm-scan))

(defun olm-summary-review ()
  "Put the review mark (`*')."
  (interactive)
  (olm-summary-insert-mark "*"))

(defun olm-summary-mark-refile ()
  (interactive)
  (let ((to (olm-pick-a-folder)))
    (goto-char (point-min))
    (while (re-search-forward "^\\*" nil t)
      (olm-summary-refile to "o"))))

;;; A helper function for olm-summary-mode functions.
(defun olm-summary-message-entry-id ()
  (let ((n (line-number-at-pos)))
    (with-current-buffer (olm-buf-entry-ids)
      (goto-line n)
      (re-search-forward "\\b[0-9A-Z]\\{140\\}\\b" nil t)
      (match-string-no-properties 0))))

(defun olm-summary-insert-mark (mark)
  (beginning-of-line)
  (setq-local buffer-read-only nil)
  (delete-char 1)
  (insert mark)
  (setq-local buffer-read-only t))

(defun olm-open-message (entry-id &optional window)
  (let ((mbuf (olm-buf-message)))
    (with-current-buffer mbuf
      (olm-do-command (format "Olm.message %S" entry-id) t)
      (let ((inhibit-read-only t))
        (olm-hide-entry-id-line)
        (olm-message-mode)
        (goto-char (point-min))))
    (if window
        (set-window-buffer window mbuf)
      (switch-to-buffer mbuf))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; olm-draft
;;;
(defvar olm-draft-mode-map nil)
(defvar olm-draft-mode-hook nil)

(unless olm-draft-mode-map
  (setq olm-draft-mode-map (make-sparse-keymap))
  (define-key olm-draft-mode-map "\C-c\C-q" 'olm-draft-kill)
  (define-key olm-draft-mode-map "\C-c\C-c" 'olm-draft-send-message)
  (define-key olm-draft-mode-map "\C-c\C-s" 'olm-draft-save-message))

(defun olm-draft-mode ()
  (interactive)
  (setq major-mode 'olm-draft-mode)
  (setq mode-name "Olm Draft")
  (setq-local line-move-ignore-invisible t)
  (font-lock-mode 1)
  (use-local-map olm-draft-mode-map)
  (run-hooks 'olm-draft-mode-hook))

(add-hook 'olm-draft-mode-hook 'olm-message-mode-keyword)

(defun olm-draft-kill ()
  (interactive)
  (kill-buffer)
  (olm-scan))

(defun olm-draft-save-message ()
  (interactive)
  (message "Olm: saving message ...")
  (olm-do-command-buf "Olm.save_message")
  (olm-draft-kill))

(defun olm-draft-send-message ()
  (interactive)
  (message "Olm: sending message ...")
  (olm-do-command-buf "Olm.send_message")
  (olm-draft-kill))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; olm-draft-reply-all
;;;
(defvar olm-draft-reply-all-map nil)
(defvar olm-draft-reply-all-hook nil)

(unless olm-draft-reply-all-map
  (setq olm-draft-reply-all-map (make-sparse-keymap))
  (define-key olm-draft-reply-all-map "\C-c\C-q" 'olm-draft-kill)
  (define-key olm-draft-reply-all-map "\C-c\C-c" 'olm-draft-reply-all-send-message)
  (define-key olm-draft-reply-all-map "\C-c\C-s" 'olm-draft-reply-all-save-message))

(defun olm-draft-reply-all-mode ()
  (interactive)
  (setq major-mode 'olm-draft-reply-all-mode)
  (setq mode-name "Olm Draft Reply All")
  (setq-local line-move-ignore-invisible t)
  (font-lock-mode 1)
  (use-local-map olm-draft-reply-all-map)
  (run-hooks 'olm-draft-reply-all-hook))

(defun olm-draft-reply-all-do-command (cmd msg)
  (message msg)
  (save-restriction
    (widen)
    (olm-do-command-buf cmd)))

(defun olm-draft-reply-all-save-message ()
  (interactive)
  (olm-draft-reply-all-do-command "Olm.update_message_body_and_save"
                                  "Olm: saving message ...")
  (olm-draft-kill))

(defun olm-draft-reply-all-send-message ()
  (interactive)
  (olm-draft-reply-all-do-command "Olm.update_message_body_and_send"
                                  "Olm: sending message ...")
  (olm-draft-kill))

(add-hook 'olm-draft-reply-all-hook 'olm-message-mode-keyword)



(provide 'olm)
;;; olm.el ends here

