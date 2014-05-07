class RemoveCompanyFromData < ActiveRecord::Migration
  def change
    remove_column :recents, :company_id, :integer
  end
end
