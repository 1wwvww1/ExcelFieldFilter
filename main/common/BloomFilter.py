# -*- coding=utf-8 -*-

from pybloom_live import ScalableBloomFilter

def construct_bloomFilter(initial_capacity=None, error_rate=None):
    if initial_capacity == None:
        return False
    if error_rate == None:
        error_rate = 0.001
    return ScalableBloomFilter(initial_capacity = initial_capacity, error_rate = error_rate)


def query_bloomFilter(content, BloomFilter):
    if(content in BloomFilter):
        return True
    return False


def insert_bloomFilter(content, BloomFilter, needQuery = False):
    isContain = False
    if(needQuery):
        isContain = query_bloomFilter(content, BloomFilter)
    if(not isContain):
        BloomFilter.add(content)
        return True
    return False