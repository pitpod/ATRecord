from django.views.generic import TemplateView
from django.http import HttpResponse

from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.pagesizes import A4, portrait
from reportlab.platypus import Table, TableStyle
from reportlab.lib.units import mm
from reportlab.lib import colors

class ReportlabView(TemplateView):
    template_name = 'pdf_sample/sample.html'

    def get(self, request, *args, **kwargs):

        response = HttpResponse(status=200, content_type='application/pdf')
        response['Content-Disposition'] = 'filename="example.pdf"'
        # response['Content-Disposition'] = 'attachment; filename="example.pdf"'

        self._create_pdf(response)
        return response

    def _create_pdf(self, response):
        # 日本語が使えるゴシック体のフォントを設定する
        font_name = 'HeiseiKakuGo-W5'
        pdfmetrics.registerFont(UnicodeCIDFont(font_name))

        # A4縦向きのpdfを作る
        size = portrait(A4)

        # pdfを描く場所を作成：pdfの原点は左上にする(bottomup=False)
        pdf_canvas = canvas.Canvas(response)
        # ヘッダー
        font_size = 24  # フォントサイズ
        pdf_canvas.setFont("HeiseiKakuGo-W5", font_size)
        pdf_canvas.drawString(93 * mm, 770, "見積書")
        font_size = 10
        pdf_canvas.setFont("HeiseiKakuGo-W5", font_size)
        pdf_canvas.drawString(
            150 * mm, 813, f"見積発行日: "
        )
        pdf_canvas.drawString(
            150 * mm,
            800,
            "xxxxxxxxxxx-xxxxxxxxxx",
        )

        # (4) 社名
        data = [
            [f"ほげほげ会社御中", ""],
            ["案件名", "ほげほげ案件"],
            ["御見積有効限：発行日より30日", ""],
        ]

        table = Table(data, colWidths=(15 * mm, 80 * mm), rowHeights=(7 * mm))
        table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), "HeiseiKakuGo-W5", 12),
                    ("LINEABOVE", (0, 1), (-1, -1), 1, colors.black),
                    ("VALIGN", (0, 0), (1, -1), "MIDDLE"),
                    ("VALIGN", (0, 1), (0, -1), "TOP"),
                ]
            )
        )
        table.wrapOn(pdf_canvas, 20 * mm, 248 * mm)
        table.drawOn(pdf_canvas, 20 * mm, 248 * mm)



        pdf_canvas.drawString(20 * mm, 238 * mm, "下記の通り御見積申し上げます")
        # (4) 社名
        data = [
            ["合計金額（消費税込）", f"1000 円"],
        ]

        table = Table(data, colWidths=(50 * mm, 60 * mm), rowHeights=(7 * mm))
        table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (1, 2), "HeiseiKakuGo-W5", 10),
                    ("BOX", (0, 0), (2, 3), 1, colors.black),
                    ("INNERGRID", (0, 0), (1, -1), 1, colors.black),
                    ("VALIGN", (0, 0), (1, 2), "MIDDLE"),
                    ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ]
            )
        )
        table.wrapOn(
            pdf_canvas,
            20 * mm,
            218 * mm,
        )
        table.drawOn(
            pdf_canvas,
            20 * mm,
            218 * mm,
        )

        # 品目
        data = [["内容", "開始月", "終了月", "単価", "数量", "金額"]]

        for idx in range(13):
            data.append([" ", " ", " ", " ", " ", ""])

        data.append([" ", " ", " ", "合計", "", f"{1000:,}"])
        data.append([" ", " ", " ", "消費税", "", f"{1000 * 0.10:,.0f}"])
        data.append([" ", " ", " ", "税込合計金額", "", f"{1000 * 1.10:,.0f}"])
        data.append(
            [" ", " ", " ", "", "", ""],
        )

        table = Table(
            data,
            colWidths=(70 * mm, 25 * mm, 25 * mm, 20 * mm, 20 * mm, 20 * mm),
            rowHeights=6 * mm,
        )
        table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), "HeiseiKakuGo-W5", 8),
                    ("BOX", (0, 0), (-1, 13), 1, colors.black),
                    ("INNERGRID", (0, 0), (-1, 13), 1, colors.black),
                    ("LINEABOVE", (3, 11), (-1, 18), 1, colors.black),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ]
            )
        )
        table.wrapOn(pdf_canvas, 17 * mm, 100 * mm)
        table.drawOn(pdf_canvas, 17 * mm, 100 * mm)

        pdf_canvas.drawString(17 * mm, 100 * mm, "<備考>")

        table = Table(
            [[""]],
            colWidths=(180 * mm),
            rowHeights=90 * mm,
        )

        table.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), "HeiseiKakuGo-W5", 8),
                    ("BOX", (0, 0), (-1, -1), 1, colors.black),
                    ("INNERGRID", (0, 0), (-1, -1), 1, colors.black),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ]
            )
        )
        table.wrapOn(pdf_canvas, 17 * mm, 5 * mm)
        table.drawOn(pdf_canvas, 17 * mm, 5 * mm)
        pdf_canvas.showPage()

        # pdfの書き出し
        pdf_canvas.save()


