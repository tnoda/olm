require 'win32ole'
require 'singleton'

module Olm
  class App
    include Singleton

    def initialize
      @app = WIN32OLE.connect("Outlook.Application")
      @ns = @app.Session
      const_load(self.class)
    end

    def default_folder
      @ns.GetDefaultFolder(OlFolderInbox)
    end

    def deleted_items_folder
      @ns.GetDefaultFolder(OlFolderDeletedItems)
    end

    def ls(folder_id = nil)
      f = folder_id ? @ns.GetFolderFromID(folder_id) : default_folder
      n = [f.Items.Count, 100].min
      s = f.Items.Count - n + 1
      t = f.Items.Count
      res = []
      s.upto(t) do |i|
        m = f.Items(i)
        unless m.Class == OlMail
          n -= 1
          next
        end
        entry_id = m.EntryID
        received_at = m.ReceivedTime.to_s.split(' ').first
        from = m.SenderName
        subject = m.Subject
        flag = m.IsMarkedAsTask ? '!' : ' '
        res << sprintf("%s %s  %-12.12s  %-20.20s %s",
          entry_id, flag, received_at, from, subject)
      end
      res.unshift(n.to_s)
    end

    def send_and_receive
      @ns.SendAndReceive(false)
    end

    def message(entry_id)
      m = @ns.GetItemFromID(entry_id)
      res = [entry_id]
      res << sprintf("From: %s", m.SenderName)
      res << sprintf("To: %s", m.To)
      res << sprintf("Cc: %s", m.CC) if m.CC.to_s.length > 0
      res << sprintf("Subject: %s", m.Subject)
      res << sprintf("ReceivedAt: %s", m.ReceivedTime)
      if m.Attachments.Count > 0
        buf = []
        m.Attachments.each do |a|
          buf << a.DisplayName
        end
        res << sprintf("Attachments: %s", buf.join("; "))
      end
      res << sprintf("---- ")
      if m.BodyFormat != OlFormatPlain
        m2 = m.Copy
        m2.BodyFormat = OlFormatPlain
        res << m2.Body.split("\r\n")
        m2.Move(deleted_items_folder)
      else
        res << m.Body.split("\r\n")
      end
      res
    end

    def toggle_task_flag(entry_id)
      m = @ns.GetItemFromID(entry_id)
      if m.IsMarkedAsTask
        m.ClearTaskFlag()
      else
        m.MarkAsTask(OlMarkNoDate)
      end
      m.Save
    end

    def mark_as_read(entry_id)
      m = @ns.GetItemFromID(entry_id)
      m.UnRead = false
      m.Save
    end

    def create_message(io)
      d = read_draft(io)
      m = @app.CreateItem(OlMailItem)
      m.BodyFormat = OlFormatPlain
      m.To = d[:to] if d[:to]
      m.CC = d[:cc] if d[:cc]
      m.BCC = d[:bcc] if d[:bcc]
      m.Subject = d[:subject] if d[:subject]
      m.Body = d[:body] if d[:body]
      m
    end

    def create_reply_all_message(entry_id)
      m = @ns.GetItemFromID(entry_id)
      r = m.ReplyAll
      r.BodyFormat = OlFormatPlain
      r.Save
      r.EntryID
    end

    def create_forward_message(entry_id)
      m = @ns.GetItemFromID(entry_id)
      r = m.Forward
      r.BodyFormat = OlFormatPlain
      r.Save
      r.EntryID
    end

    def update_message_body(io)
      d = read_draft(io)
      m = @ns.GetItemFromID(d[:entry_id])
      m.BodyFormat = OlFormatPlain
      m.Body = d[:body]
      m.BCC = d[:bcc] if d[:bcc]
      m
    end

    def update_forward_message_body(io)
      d = read_draft(io)
      m = @ns.GetItemFromID(d[:entry_id])
      m.BodyFormat = OlFormatPlain
      m.Body = d[:body]
      m.To = d[:to] if d[:to]
      m.BCC = d[:bcc] if d[:bcc]
      m
    end

    def save_attachments(entry_id, path)
      m = @ns.GetItemFromID(entry_id)
      m.Attachments.each do |a|
        a.SaveAsFile(path + a.DisplayName)
      end
    end

    def execute_refile(io)
      io.each_line do |line|
        line.chomp!
        next unless /^(\h+) (\h+)/ =~ line
        move($1, $2)
      end
    end

    private

    def read_draft(io)
      {}.update(entry_id: io.readline.chomp)
        .update(read_draft_headers(io))
        .update(body: io.readlines.map { |s| s.chomp }.join("\r\n"))
    end

    def read_draft_headers(io)
      headers = {}
      io.each_line do |line|
        break if /^---- / =~ line
        line.chomp!
        next unless /^([^:]+): (.*)/ =~ line
        headers[$1.downcase.intern] = $2
      end
      headers
    end

    def const_load(klass)
      WIN32OLE.const_load(@app, klass)
    end

    def move(from, to)
      item = @ns.GetItemFromID(from)
      folder = @ns.GetFolderFromID(to)
      item.Move(folder)
    end
  end
end
