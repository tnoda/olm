require 'olm/app'

module Olm
  VERSION = '0.0.1'

  extend self

  def ls(folder_id = nil)
    puts Olm::App.instance.ls(folder_id)
  end

  def send_and_receive
    Olm::App.instance.send_and_receive
  end

  def message(entry_id)
    puts Olm::App.instance.message(entry_id)
  end

  def toggle_task_flag(entry_id)
    Olm::App.instance.toggle_task_flag(entry_id)
  end

  def mark_as_read(entry_id)
    Olm::App.instance.mark_as_read(entry_id)
  end

  def save_message
    Olm::App.instance.create_message(ARGF).Save
  end

  def send_message
    Olm::App.instance.create_message(ARGF).Send
  end
end
