from twocaptcha import TwoCaptcha


def solve_captcha(img_path):

    config = {
                'apiKey': '5c8194a3f88ea370c8d73c89f2e71708',
                'pollingInterval': 5
            }

    solver = TwoCaptcha(**config)

    try:
        result = solver.normal(img_path)

    except Exception as e:
        return e

    else:
        return result['code']