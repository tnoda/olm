(require 'dash)

(defvar olm-folder-id nil)
(defvar olm-default-bcc nil)
(defvar olm-attachment-path nil)

(defun olm
  ()
  (interactive)
  (olm-scan))

(defun olm-scan
  ()
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
         (ibuf (let ((buf (olm-buf-entry-ids)))
                 (with-current-buffer buf
                   (erase-buffer))
                 buf)))
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
(defun olm-ls
  ()
  (with-current-buffer (olm-buf-ls)
    (setq-local buffer-read-only nil)
    (erase-buffer)
    (let ((command (if olm-folder-id
                       (format "Olm.ls %S" olm-folder-id)
                     "Olm.ls")))
      (olm-do-command command t))))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; gereral helper functions
;;; 
(defun olm-do-command
  (command &optional buf)
  (let ((buf (or buf (get-buffer-create "*Messages*"))))
    (call-process "ruby" nil buf nil
                "-r" "rubygems" "-r" "olm" "-e" command)))

(defun olm-sync
  ()
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

(defun olm-hide-entry-id-line
  ()
  (interactive)
  (let ((pos (point)))
    (narrow-to-region (progn
                        (goto-line 2)
                        (point))
                      (point-max))
    (goto-char pos)))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; buffers
;;; 
(defun olm-buf-ls
  ()
  (get-buffer-create "*olm-ls*"))

(defun olm-buf-summary
  ()
  (get-buffer-create "*olm-summary*"))

(defun olm-buf-entry-ids
  ()
  (get-buffer-create "*olm-entry-ids*"))

(defun olm-buf-message
  ()
  (let ((buf (get-buffer-create "*olm-message*")))
    (with-current-buffer buf
      (setq-local buffer-read-only nil)
      (erase-buffer))
    buf))

(defun olm-buf-draft
  ()
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

(defun olm-buf-draft-reply-all
  ()
  (let ((buf (get-buffer-create "*olm-draft-reply-all*")))
    (with-current-buffer buf
      (setq-local buffer-read-only nil)
      (erase-buffer))
    buf))


;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;
;;;
;;; olm-message
;;; 
(defvar olm-message-mode-map nil)
(defvar olm-message-mode-hook nil)

(defun olm-message-mode
  ()
  (interactive)
  (setq major-mode 'olm-summary-mode)
  (setq mode-name "Olm Message")
  (font-lock-mode 1)
  (setq-local buffer-read-only t)
  (setq-local line-move-ignore-invisible t)
  (run-hooks 'olm-message-mode-hook))

(defun olm-message-mode-keyword
  ()
  (font-lock-add-keywords
   nil
   '(("^From:" . font-lock-keyword-face)
     ("^To:" . font-lock-keyword-face)
     ("^Cc:" . font-lock-keyword-face)
     ("^Subject:" . font-lock-keyword-face)
     ("^ReceivedAt:" . font-lock-keyword-face)
     ("^> .*$" . font-lock-comment-face)
     ("^From: \\(.*\\)$" 1 font-lock-warning-face)
     ("^To: \\(.*\\)$" 1 font-lock-negation-char-face)
     ("^Cc: \\(.*\\)$" 1 font-lock-constant-face)
     ("^Subject: \\(.*\\)$" 1 font-lock-variable-name-face)
     ("^ReceivedAt: \\(.*\\)$" 1 font-lock-type-face))))

(add-hook 'olm-message-mode-hook 'olm-message-mode-keyword)

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
  (define-key olm-summary-mode-map [backspace] 'olm-summary-scroll-message-backward)
  (define-key olm-summary-mode-map "p" 'olm-summary-display-up)
  (define-key olm-summary-mode-map "n" 'olm-summary-display-down)
  (define-key olm-summary-mode-map "!" 'olm-summary-toggle-flag)
  (define-key olm-summary-mode-map "w" 'olm-summary-write)
  (define-key olm-summary-mode-map "A" 'olm-summary-reply-all)
  (define-key olm-summary-mode-map "\C-c\C-a" 'olm-summary-save-attachments))

(defun olm-summary-inc
  ()
  (interactive)
  (olm-sync)
  (olm-scan))

(defun olm-summary-quit
  ()
  (interactive)
  (delete-other-windows-vertically)
  (quit-window))

(defun olm-summary-mode
  ()
  (interactive)
  (use-local-map olm-summary-mode-map)
  (setq major-mode 'olm-summary-mode)
  (setq mode-name "Olm Summary")
  (setq-local buffer-read-only t)
  (setq-local truncate-lines t)
  (setq-local line-move-ignore-invisible t)
  (run-hooks 'olm-summary-mode-hook))

(defun olm-summary-open-message
  ()
  (interactive)
  (delete-other-windows-vertically)
  (let* ((ln0 (line-number-at-pos))
         (msg-window (split-window-below 10))
         (mbuf (olm-buf-message))
         (entry-id (progn
                     (goto-line ln0)
                     (olm-mail-item-entry-id-at))))
    (with-current-buffer mbuf
      (setq-local buffer-file-coding-system 'utf-8-dos))
    (olm-do-command (format "Olm.message %S" entry-id) mbuf)
    (with-current-buffer mbuf
      (let ((inhibit-read-only t))
        (put-text-property (progn
                             (goto-char (point-min))
                             (point))
                           (line-end-position)
                           'read-only
                           nil))
      (olm-hide-entry-id-line)
      (olm-message-mode)
      (goto-char (point-min)))
    (set-window-buffer msg-window mbuf)
    (olm-summary-mark-message-as-read)))

(defun olm-summary-mark-message-as-read
  ()
  (interactive)
  (olm-do-command (format "Olm.mark_as_read %S" (olm-mail-item-entry-id-at))))

(defun olm-summary-scroll-message-forward
  ()
  (interactive)
  (recenter)
  (scroll-other-window olm-summary-scroll-lines))

(defun olm-summary-scroll-message-backward
  ()
  (interactive)
  (recenter)
  (scroll-other-window (- olm-summary-scroll-lines)))

(defun olm-summary-display-up
  ()
  (interactive)
  (forward-line -1)
  (recenter)
  (olm-summary-open-message))

(defun olm-summary-display-down
  ()
  (interactive)
  (forward-line 1)
  (recenter)
  (olm-summary-open-message))

(defun olm-summary-toggle-flag
  ()
  (interactive)
  (let ((n (line-number-at-pos)))
    (olm-do-command (format "Olm.toggle_task_flag %S"
                            (olm-mail-item-entry-id-at)))
    (olm-scan)
    (goto-line n)))

(defun olm-summary-write
  ()
  (interactive)
  (delete-other-windows-vertically)
  (let ((buf (olm-buf-draft)))
    (with-current-buffer buf
      (olm-hide-entry-id-line)
      (olm-draft-mode))
    (switch-to-buffer buf)))

(defun olm-summary-reply-all
  ()
  (interactive)
  (let ((entry-id (olm-mail-item-entry-id-at))
        (buf (olm-buf-draft-reply-all)))
    (olm-do-command (format "Olm.create_reply_all_message %S" entry-id)
                    buf)
    (with-current-buffer buf
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
    (delete-other-windows-vertically)
    (switch-to-buffer buf)))

(defun olm-summary-save-attachments
  ()
  (interactive)
  (olm-do-command (format "Olm.save_attachments(%S, %S)"
                          (olm-mail-item-entry-id-at)
                          olm-attachment-path)))

;;; A helper function for olm-summary-mode functions.
(defun olm-mail-item-entry-id-at
  ()
  (interactive)
  (let ((n (line-number-at-pos)))
    (with-current-buffer (olm-buf-entry-ids)
      (goto-line n)
      (buffer-substring (line-beginning-position) (line-end-position)))))


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

(defun olm-draft-mode
  ()
  (interactive)
  (setq major-mode 'olm-draft-mode)
  (setq mode-name "Olm Draft")
  (setq-local line-move-ignore-invisible t)
  (font-lock-mode 1)
  (use-local-map olm-draft-mode-map)
  (run-hooks 'olm-draft-mode-hook))

(add-hook 'olm-draft-mode-hook 'olm-message-mode-keyword)

(defun olm-draft-kill
  ()
  (interactive)
  (kill-buffer)
  (olm-scan))

(defun olm-draft-do-command
  (command)
  (call-process-region (point-min) (point-max)
                       "ruby" nil (get-buffer-create "*Messages*") nil
                       "-r" "rubygems" "-r" "olm" "-e" command))

(defun olm-draft-save-message
  ()
  (interactive)
  (message "Olm: saving message ...")
  (olm-draft-do-command "Olm.save_message")
  (olm-draft-kill))

(defun olm-draft-send-message
  ()
  (interactive)
  (message "Olm: sending message ...")
  (olm-draft-do-command "Olm.send_message")
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

(defun olm-draft-reply-all-mode
  ()
  (interactive)
  (setq major-mode 'olm-draft-reply-all-mode)
  (setq mode-name "Olm Draft Reply All")
  (setq-local line-move-ignore-invisible t)
  (font-lock-mode 1)
  (use-local-map olm-draft-reply-all-map)
  (run-hooks 'olm-draft-reply-all-hook))

(defun olm-draft-reply-all-do-command
  (cmd msg)
  (message msg)
  (save-restriction
    (widen)
    (olm-draft-do-command cmd))
  (olm-draft-kill))

(defun olm-draft-reply-all-save-message
  ()
  (interactive)
  (olm-draft-reply-all-do-command "Olm.update_message_body_and_save"
                                  "Olm: saving message ..."))

(defun olm-draft-reply-all-send-message
  ()
  (interactive)
  (olm-draft-reply-all-do-command "Olm.update_message_body_and_send"
                                  "Olm: sending message ..."))

(add-hook 'olm-draft-reply-all-hook 'olm-message-mode-keyword)
