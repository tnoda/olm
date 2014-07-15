require 'win32ole'
require 'singleton'
require 'nkf'

module Olm
  class App
    include Singleton

    def initialize
      @app = WIN32OLE.connect("Outlook.Application")
      @ns = @app.Session
      @enc = @ns.Folders.GetFirst.Name.encoding
      const_load(self.class)
    end

    def ls(folder_id = nil)
      f = folder_id ? @ns.GetFolderFromID(folder_id) : default_folder
      res = "#{f.Items.Count}\n"
      f.Items.each do |m|
        entry_id = m.EntryID
        received_at = m.ReceivedTime.to_s.split(' ').first
        from = m.SenderName
        subject = m.Subject
        flag = m.IsMarkedAsTask ? '!' : ' '
        res << sprintf("%s %s  %-12.12s  %-20.20s %s\n",
          entry_id, flag, received_at, u(from), u(subject))
      end
      res
    end

    def send_and_receive
      @ns.SendAndReceive(false)
    end

    def message(entry_id)
      m = @ns.GetItemFromID(entry_id)
      res = entry_id + "\n"
      res << sprintf("From: %s\n", m.SenderName)
      res << sprintf("To: %s\n", m.To)
      res << sprintf("Cc: %s\n", m.CC) if m.CC.to_s.length > 0
      res << sprintf("Subject: %s\n", m.Subject)
      res << sprintf("ReceivedAt: %s\n", m.ReceivedTime)
      if m.Attachments.Count > 0
        buf = []
        m.Attachments.each do |a|
          buf << a.DisplayName
        end
        res << sprintf("Attachments: %s\n", buf.join("; "))
      end
      res << sprintf("---- \n")
      if m.BodyFormat != OlFormatPlain
        m2 = m.Copy
        m2.BodyFormat = OlFormatPlain
        res << m2.Body
        m2.Move(@ns.GetDefaultFolder(OlFolderDeletedItems))
      else
        res << m.Body
      end
      NKF.nkf('-w -Lu', res)
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
      x = {:body => ''}
      header = true
      io.set_encoding(Encoding::UTF_8).each_line do |line|
        line.chomp!
        if header
          if /^---- / =~ line
            header = false
          else
            next unless /^(.*?): (.*)/ =~ line
            x[$1] = $2
          end
        else
          x[:body] << line
          x[:body] << "\r\n"
        end
      end
      m = @app.CreateItem(OlMailItem)
      m.BodyFormat = OlFormatPlain
      m.To = x['To'] if x['To']
      m.CC = x['Cc'] if x['Cc']
      m.BCC = x['Bcc'] if x['Bcc']
      m.Subject = w(x['Subject']) if x['Subject']
      m.Body = w(x[:body]) if x[:body]
      m
    end

    def create_reply_all_message(entry_id)
      m = @ns.GetItemFromID(entry_id)
      r = m.ReplyAll
      r.BodyFormat = OlFormatPlain
      r.Save
      r.EntryID
    end

    def update_message_body(io)
      entry_id = nil
      header = true
      bcc = nil
      body = ''
      io.set_encoding(Encoding::UTF_8).each_line do |line|
        line.chomp!
        if entry_id.nil?
          entry_id = line
        elsif header
          bcc = $1 if /^Bcc: (.*)/ =~ line
          header = false if /^---- / =~ line
        else
          body << line
          body << "\n"
        end
      end
      m = @ns.GetItemFromID(entry_id)
      m.BodyFormat = OlFormatPlain
      m.Body = NKF.nkf('-s -Lw', body)
      m.BCC = bcc if bcc
      m
    end

    def save_attachments(entry_id, path)
      m = @ns.GetItemFromID(entry_id)
      m.Attachments.each do |a|
        a.SaveAsFile(w(path + a.DisplayName))
      end
    end


    private

    def const_load(klass)
      WIN32OLE.const_load(@app, klass)
    end

    def default_folder
      @ns.GetDefaultFolder(OlFolderInbox)
    end

    def u(str)
      str.encode(Encoding::UTF_8)
    end

    def w(str)
      str.encode(@enc)
    end
  end
end
