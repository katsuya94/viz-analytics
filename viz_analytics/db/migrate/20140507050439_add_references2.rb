class AddReferences2 < ActiveRecord::Migration
  def change
  	add_reference :data, :previous
  end
end
