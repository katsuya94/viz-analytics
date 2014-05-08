class DropRecent < ActiveRecord::Migration
  def change
  	drop_table :recents
  end
end
