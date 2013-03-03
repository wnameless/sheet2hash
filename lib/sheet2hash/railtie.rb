module Sheet2hash
  class Railtie < Rails::Railtie
    initializer 'sheet2hash' do
      ActionController::Base.send :include, Sheet2hash
    end
  end
end
