class CreateRecentsOld < ActiveRecord::Migration
  def change
    create_table :recents do |t|

      t.timestamps
    end
    add_reference :recents, :company
    add_reference :recents, :datum
    add_reference :recents, :metric
  end
end
