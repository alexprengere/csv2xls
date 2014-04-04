#!/usr/bin/env python

"""
Random CSV generator.
"""

if __name__ == '__main__':

    for name in range(2):

        with open('test%s.csv' % name, 'w') as f:
            for i in range(65536):
                f.write(','.join('0123456789') + '\n')
