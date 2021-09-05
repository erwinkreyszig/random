local deployment instructions

requirements
  1 . docker installed
steps
  1. pull image from registry
    docker pull erwinkreyszig/twitter-demo:latest
  2. run container
    docker run --name=twitter-demo -d -p 8080:5000 erwinkreyszig/twitter-demo:latest 
  3. issue requests
    curl localhost:8080/hashtags/whatever-hashtag-here?limit=2
    curl localhost:8080/users/whatever-user-here?limit=2
