# -*- coding: utf-8 -*-
from datetime import datetime
from enum import Enum
from flask import current_app
from requests import get
from urllib.parse import quote, urlencode


class Endpoints(Enum):
    TWEET = 'tweet'
    USER = 'user'
    USER_TWEETS = 'user-tweets'


BASE_URL = 'https://api.twitter.com/2'
TWITTER_ENDPOINTS = {
    Endpoints.TWEET: '/tweets/search/recent',
    Endpoints.USER: '/users/by/username/{param}',
    Endpoints.USER_TWEETS: '/users/{param}/tweets'
}
TWEET_COUNT_LOWER_LIMIT = 10
TWEET_COUNT_UPPER_LIMIT = 500
RETURNED_TWEET_COUNT_DEFAULT = 30
TWEET_FIELDS = ['author_id', 'created_at', 'public_metrics']
USER_FIELDS = ['id', 'username', 'name', 'url']
EXPANSIONS = ['author_id']


def create_headers():
    """Generates the request header for twitter api."""
    return {'Authorization': f'Bearer {current_app.config["BEARER_TOKEN"]}'}


def create_url(endpoint, param=None):
    """Generates twitter api endpoint to access."""
    if param:
        return f'{BASE_URL}{TWITTER_ENDPOINTS[endpoint].format(param=param)}'
    return f'{BASE_URL}{TWITTER_ENDPOINTS[endpoint]}'


def create_query(limit, query=None):
    """Generates the common query fields for a tweet query."""
    fields = {'tweet.fields': ','.join(TWEET_FIELDS),
              'user.fields': ','.join(USER_FIELDS),
              'expansions': ','.join(EXPANSIONS)}
    if query is not None:
        fields.update({'query': query})
    if TWEET_COUNT_LOWER_LIMIT <= limit <= TWEET_COUNT_UPPER_LIMIT:
        fields.update({'max_results': limit})
    return urlencode(fields)


def fetch_tweets_by_hashtag(hashtag, limit=RETURNED_TWEET_COUNT_DEFAULT):
    """Fetches tweets that have the specified hashtag.
    Returned tweet count is specified by the limit parameter (default is 30).
    """
    query = create_query(limit, f'#{hashtag}')
    url = f'{create_url(Endpoints.TWEET)}?{query}'
    print(url)
    results = get(url, headers=create_headers())
    # TODO: add error handling here later (when results.staus_code != 200)
    print(results.status_code)
    tweets = results.json()['data']
    if limit < TWEET_COUNT_LOWER_LIMIT:
        return format_results(tweets[:limit])
    return format_results(tweets)


def fetch_tweets_by_user(user, limit=RETURNED_TWEET_COUNT_DEFAULT):
    """Fetches tweets by a specified user.
    Returned tweet count is specified by the limit parameter (default is 30).
    """
    # fetch user's id
    url = create_url(Endpoints.USER, user)
    result = get(url, headers=create_headers())
    # TODO: add error handling here later (when results.staus_code != 200)
    user_id = result.json()['data']['id']
    # fetch tweets with user_id
    query = create_query(limit)
    url = f'{create_url(Endpoints.USER_TWEETS, user_id)}?{query}'
    results = get(url, headers=create_headers())
    tweets = results.json()['data']
    if limit < TWEET_COUNT_LOWER_LIMIT:
        return format_results(tweets[:limit])
    return format_results(tweets)


def format_results(data):
    """Formats the tweet data to the desired structure."""
    formatted_results = []
    for item in data:
        _tmp = {
            # maybe get user data from TWITTER_ENDPOINTS['user'] endpoint
            # then add it here?
            'account': {
                'fullname': '',  # todo
                'href': '',  # todo
                'id': item['author_id']
            }
        }
        _tmp['hashtags'] = []  # todo
        _tmp['text'] = item['text']
        public_metrics = item['public_metrics']
        _tmp['likes'] = public_metrics['like_count']
        _tmp['retweets'] = public_metrics['retweet_count']
        _tmp['replies'] = public_metrics['reply_count']
        _tmp['quotes'] = public_metrics['quote_count']
        t = datetime.strptime(item['created_at'], '%Y-%m-%dT%H:%M:%S.%fZ')
        _tmp['date'] = t.strftime('%I:%S %p - %d %b %Y')
        formatted_results.append(_tmp)
    return formatted_results


if __name__ == '__main__':
    # TODO: move these to a separate test file
    # test for format_results method
    dummy_data = [{'author_id': '1',
                   'created_at': '2021-09-05T13:30:00.000Z',
                   'id': '1',
                   'public_metrics': {'like_count': 1,
                                      'quote_count': 1,
                                      'reply_count': 1,
                                      'retweet_count': 1},
                   'text': 'dummy'},
                  {'author_id': '2',
                   'created_at': '2021-09-05T13:40:00.000Z',
                   'id': '2',
                   'public_metrics': {'like_count': 2,
                                      'quote_count': 2,
                                      'reply_count': 2,
                                      'retweet_count': 2},
                   'text': 'dummy'}]
    results = format_results(dummy_data)
    assert len(dummy_data) == len(results)
    for i, result in enumerate(results):
        _tmp = dummy_data[i]
        assert result['account']['id'] == _tmp['author_id']
        assert result['text'] == _tmp['text']
        assert result['likes'] == _tmp['public_metrics']['like_count']
        assert result['retweets'] == _tmp['public_metrics']['retweet_count']
        assert result['replies'] == _tmp['public_metrics']['reply_count']
        assert result['quotes'] == _tmp['public_metrics']['quote_count']
        assert result['date'] == datetime.strptime(_tmp['created_at'],
                                                   '%Y-%m-%dT%H:%M:%S.%fZ') \
            .strftime('%I:%S %p - %d %b %Y')
    # test for create_query
    assert create_query(3) == 'tweet.fields=author_id%2Ccreated_at%2Cpublic_metrics&user.fields=id%2Cusername%2Cname%2Curl&expansions=author_id'
    assert create_query(3, '#dummy') == 'tweet.fields=author_id%2Ccreated_at%2Cpublic_metrics&user.fields=id%2Cusername%2Cname%2Curl&expansions=author_id&query=%23dummy'
    assert create_query(11) == 'tweet.fields=author_id%2Ccreated_at%2Cpublic_metrics&user.fields=id%2Cusername%2Cname%2Curl&expansions=author_id&max_results=11'
    assert create_query(11, '#dummy') == 'tweet.fields=author_id%2Ccreated_at%2Cpublic_metrics&user.fields=id%2Cusername%2Cname%2Curl&expansions=author_id&query=%23dummy&max_results=11'
    assert create_query(501) == 'tweet.fields=author_id%2Ccreated_at%2Cpublic_metrics&user.fields=id%2Cusername%2Cname%2Curl&expansions=author_id'
    assert create_query(501, '#dummy') == 'tweet.fields=author_id%2Ccreated_at%2Cpublic_metrics&user.fields=id%2Cusername%2Cname%2Curl&expansions=author_id&query=%23dummy'
