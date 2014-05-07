class AddReferences < ActiveRecord::Migration
  def change
  	add_reference :data, :metric
  end
end
