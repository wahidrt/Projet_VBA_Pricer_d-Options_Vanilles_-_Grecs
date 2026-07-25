import math
import unittest


def normal_cdf(x: float) -> float:
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def normal_pdf(x: float) -> float:
    return math.exp(-0.5 * x * x) / math.sqrt(2.0 * math.pi)


def d1(spot: float, strike: float, rate: float, sigma: float,
       maturity: float, dividend: float = 0.0) -> float:
    return (
        math.log(spot / strike)
        + (rate - dividend + 0.5 * sigma * sigma) * maturity
    ) / (sigma * math.sqrt(maturity))


def call_price(spot: float, strike: float, rate: float, sigma: float,
               maturity: float, dividend: float = 0.0) -> float:
    first = d1(spot, strike, rate, sigma, maturity, dividend)
    second = first - sigma * math.sqrt(maturity)
    return (
        spot * math.exp(-dividend * maturity) * normal_cdf(first)
        - strike * math.exp(-rate * maturity) * normal_cdf(second)
    )


def put_price(spot: float, strike: float, rate: float, sigma: float,
              maturity: float, dividend: float = 0.0) -> float:
    first = d1(spot, strike, rate, sigma, maturity, dividend)
    second = first - sigma * math.sqrt(maturity)
    return (
        strike * math.exp(-rate * maturity) * normal_cdf(-second)
        - spot * math.exp(-dividend * maturity) * normal_cdf(-first)
    )


def call_delta(spot: float, strike: float, rate: float, sigma: float,
               maturity: float, dividend: float = 0.0) -> float:
    return (
        math.exp(-dividend * maturity)
        * normal_cdf(d1(spot, strike, rate, sigma, maturity, dividend))
    )


def gamma(spot: float, strike: float, rate: float, sigma: float,
          maturity: float, dividend: float = 0.0) -> float:
    return (
        math.exp(-dividend * maturity)
        * normal_pdf(d1(spot, strike, rate, sigma, maturity, dividend))
        / (spot * sigma * math.sqrt(maturity))
    )


class BlackScholesReferenceTests(unittest.TestCase):
    def test_put_call_parity_with_dividend(self) -> None:
        spot = 100.0
        strike = 105.0
        rate = 0.03
        dividend = 0.012
        sigma = 0.24
        maturity = 0.75

        call = call_price(spot, strike, rate, sigma, maturity, dividend)
        put = put_price(spot, strike, rate, sigma, maturity, dividend)
        expected = (
            spot * math.exp(-dividend * maturity)
            - strike * math.exp(-rate * maturity)
        )
        self.assertAlmostEqual(call - put, expected, places=12)

    def test_delta_and_gamma_match_finite_differences(self) -> None:
        spot = 100.0
        strike = 100.0
        rate = 0.02
        sigma = 0.25
        maturity = 30.0 / 365.0
        step = 0.01

        center = call_price(spot, strike, rate, sigma, maturity)
        up = call_price(spot + step, strike, rate, sigma, maturity)
        down = call_price(spot - step, strike, rate, sigma, maturity)

        numerical_delta = (up - down) / (2.0 * step)
        numerical_gamma = (up - 2.0 * center + down) / (step * step)

        self.assertAlmostEqual(
            call_delta(spot, strike, rate, sigma, maturity),
            numerical_delta,
            places=7,
        )
        self.assertAlmostEqual(
            gamma(spot, strike, rate, sigma, maturity),
            numerical_gamma,
            places=6,
        )

    def test_current_portfolio_does_not_accumulate_between_spots(self) -> None:
        rate = 0.02
        sigma = 0.25
        maturity = 30.0 / 365.0
        initial_call = call_price(100.0, 100.0, rate, sigma, maturity)
        positions = [
            (1.0, True, 100.0, initial_call),
            (-1.0, True, 105.0, 2.0),
            (1.0, False, 120.0, 5.0),
        ]

        def portfolio_value(spot: float) -> float:
            total = 0.0
            for sign, is_call, strike, _premium in positions:
                price = (
                    call_price(spot, strike, rate, sigma, maturity)
                    if is_call
                    else put_price(spot, strike, rate, sigma, maturity)
                )
                total += sign * price
            return total

        self.assertAlmostEqual(portfolio_value(50.0), 69.80290176895662, places=10)
        self.assertAlmostEqual(portfolio_value(55.0), 64.80290176895662, places=10)
        self.assertAlmostEqual(portfolio_value(100.0), 21.63180112062848, places=10)
        self.assertAlmostEqual(portfolio_value(500.0), 4.991787573706517, places=10)


if __name__ == "__main__":
    unittest.main()
