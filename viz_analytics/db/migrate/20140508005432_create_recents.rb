class CreateRecents < ActiveRecord::Migration
  def change
    create_table :recents do |t|
      t.timestamps
    end
    add_reference :recents, :metric
    add_reference :recents, :datum
  end
end
