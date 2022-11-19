import wrapt_timeout_decorator

t = 1
# デコレータでタイムアウトの秒数を設定
@wrapt_timeout_decorator.timeout(dec_timeout=t)
def func():
    while True:
        pass


def callback():
    print("timeout")


if __name__ == "__main__":
    try:
        func()
    except TimeoutError:
        callback()
