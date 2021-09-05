local deployment instructions

requirements:<br />1. docker installed

steps:<br />1. pull image from registry<br />
docker pull erwinkreyszig/twitter-demo:latest<br />
2. run container<br />
docker run --name=twitter-demo -d -p 8080:5000 erwinkreyszig/twitter-demo:latest <br />
3. issue requests<br />
curl localhost:8080/hashtags/whatever-hashtag-here?limit=2<br />
curl localhost:8080/users/whatever-user-here?limit=2
