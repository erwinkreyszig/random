# -*- coding: utf-8 -*-
from flask import Flask, current_app, jsonify, request
import os

from lib.fetch_tweets import RETURNED_TWEET_COUNT_DEFAULT


app = Flask(__name__)
with app.app_context():
    current_app.config['BEARER_TOKEN'] = os.environ['BEARER_TOKEN']


def get_limit_param(request):
    """Helper method for getting the passed limit parameter."""
    try:
        limit = request.args.get('limit', None)
        limit = int(limit) if limit is not None else RETURNED_TWEET_COUNT_DEFAULT
    except ValueError:
        limit = False
    return limit


@app.route('/')
def index():
    """Default endpoint."""
    return jsonify({'response': 200})


@app.route('/hashtags/<hashtag>', methods=['GET'])
def get_tweets_with_hashtag(hashtag):
    """Endpoint to fetch tweets that have a specified hashtag."""
    from lib.fetch_tweets import fetch_tweets_by_hashtag
    print(hashtag)
    print(request.args)
    limit = get_limit_param(request)
    if not limit:
        # TODO: add error handling here, but for now return an empty list
        return jsonify([])
    tweets = fetch_tweets_by_hashtag(hashtag, limit)
    return jsonify(tweets)


@app.route('/users/<user>', methods=['GET'])
def get_tweets_by_user(user):
    """Endpoint to fetch tweets by a specified user."""
    from lib.fetch_tweets import fetch_tweets_by_user
    limit = get_limit_param(request)
    if not limit:
        # TODO: add error handling here, but for now return an empty list
        return jsonify([])
    tweets = fetch_tweets_by_user(user, limit)
    return jsonify(tweets)
