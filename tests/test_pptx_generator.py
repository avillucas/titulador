import unittest
import os
import pptx
from src.models import TituloData
from src.models import RenderConfig
from src.pptx_generator import PPTXGenerator, adaptive_font_size, format_certificacion_and_resolucion

class TestPPTXGenerator(unittest.TestCase):
    def test_dash_padding_in_shapes(self):
        data = TituloData(
            apellido_nombre='VILLALBA JORGE HUMBERTO',
            documento='17.924.873',
            horas_cursada='540',
            titulo_linea1='Electricista en Inmuebles',
            titulo_linea2='Electricista en Inmuebles',
            resolucion='Resolución Nro. RESFC-2019-6908-GDEBA-DGCYE / RESOL-2022-2139-APN-ME--',
            emision_dia='09',
            emision_mes='Julio',
            emision_ano='26',
            modulos=['Módulo 1', 'Módulo 2'],
            fecha_egreso='09 de Junio de 2026',
            numero_egresado='1',
            numero_cpf='412',
            distrito_cpf='Lomas de Zamora'
        )

        gen = PPTXGenerator()
        out_path = 'output/test_dash_padding.pptx'
        gen.generate(data, out_path)

        self.assertTrue(os.path.exists(out_path))

        prs = pptx.Presentation(out_path)
        slide1 = prs.slides[0]

        shape_dict = {shape.shape_id: shape.text_frame.text for shape in slide1.shapes if shape.has_text_frame}

        # Shape 88: Title line 1 before hours should be padded with hyphens
        self.assertIn('Electricista en Inmuebles-', shape_dict[88])
        self.assertTrue(shape_dict[88].endswith('-'))

        # Shape 89 starts the certification; the resolution continues on shape 90.
        self.assertEqual(shape_dict[89], 'Electricista en Inmuebles')

        # Shape 90: Resolution line should end with hyphens
        self.assertIn('según Resolución Nro.', shape_dict[90])
        self.assertTrue(shape_dict[90].rstrip().endswith('---'))

        # Clean up test output
        if os.path.exists(out_path):
            os.remove(out_path)

    def test_second_line_never_contains_line_breaks(self):
        _, second_line = format_certificacion_and_resolucion(
            'Operadora/or de Carpintería y\nFabricación de Mobiliario',
            'Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE--\r\n',
        )

        self.assertNotIn('\n', second_line)
        self.assertNotIn('\r', second_line)

    def test_font_size_decreases_as_text_gets_longer(self):
        sizes = [
            adaptive_font_size('texto ' * length, preferred_size=12.0, max_chars=62)
            for length in (5, 15, 30, 60)
        ]

        self.assertEqual(sizes[0], 12.0)
        self.assertGreater(sizes[0], sizes[1])
        self.assertGreater(sizes[1], sizes[2])
        self.assertGreaterEqual(sizes[2], sizes[3])
        self.assertEqual(sizes[3], 6.5)

    def test_render_config_overrides_text_and_font_sizes(self):
        data = TituloData(titulo_linea1='Trayecto original', titulo_linea2='Carrera original')
        config = RenderConfig(
            trayecto_text='Trayecto configurable',
            carrera_line1_text='Carrera configurable',
            carrera_line2_text='Resolución configurable',
            trayecto_font_size=8.0,
            carrera_font_size=9.0,
        )
        out_path = 'output/test_render_config.pptx'
        PPTXGenerator().generate(data, out_path, config)
        prs = pptx.Presentation(out_path)
        shape_dict = {shape.shape_id: shape for shape in prs.slides[0].shapes if shape.has_text_frame}

        self.assertIn('Trayecto configurable', shape_dict[88].text)
        self.assertEqual(shape_dict[89].text, 'Carrera configurable')
        self.assertEqual(shape_dict[90].text, 'Resolución configurable')
        self.assertEqual(shape_dict[88].text_frame.paragraphs[0].font.size.pt, 8.0)
        self.assertEqual(shape_dict[89].text_frame.paragraphs[0].font.size.pt, 9.0)

        if os.path.exists(out_path):
            os.remove(out_path)

    def test_default_front_font_size_is_12_points(self):
        data = TituloData(titulo_linea1='Trayecto breve', titulo_linea2='Carrera breve')
        out_path = 'output/test_default_font_size.pptx'
        PPTXGenerator().generate(data, out_path)

        try:
            prs = pptx.Presentation(out_path)
            shape_dict = {shape.shape_id: shape for shape in prs.slides[0].shapes if shape.has_text_frame}
            self.assertEqual(shape_dict[89].text_frame.paragraphs[0].font.size.pt, 12.0)
            self.assertEqual(adaptive_font_size('Trayecto breve', 12.0, 42), 12.0)
        finally:
            if os.path.exists(out_path):
                os.remove(out_path)

    def test_multi_line_title_no_dashes_in_middle(self):
        data = TituloData(
            apellido_nombre='PEREZ JUAN MANUEL',
            documento='35.140.353',
            horas_cursada='230',
            titulo_linea1='Operadora/or de Carpintería y Fabricación de Mobiliario',
            titulo_linea2='Operadora/or de Carpintería y Fabricación de Mobiliario',
            resolucion='Resolución Nro. RESOC-2022-2450-GDEBA-DGCYE--',
            emision_dia='25',
            emision_mes='Junio',
            emision_ano='26',
            modulos=['Módulo 1'],
            fecha_egreso='25 de Mayo de 2026',
            numero_egresado='2',
            numero_cpf='412',
            distrito_cpf='Lomas de Zamora'
        )

        gen = PPTXGenerator()
        out_path = 'output/test_multi_line_dash.pptx'
        gen.generate(data, out_path)

        prs = pptx.Presentation(out_path)
        slide1 = prs.slides[0]

        shape_dict = {shape.shape_id: shape.text_frame.text for shape in slide1.shapes if shape.has_text_frame}

        # Shape 89 should NOT end with hyphens since the title continues on Shape 90
        self.assertFalse(shape_dict[89].endswith('-'))
        self.assertEqual(shape_dict[89], 'Operadora/or de Carpintería y')

        if os.path.exists(out_path):
            os.remove(out_path)

    def test_long_trayecto_name_ends_with_dashes(self):
        data = TituloData(
            titulo_linea1='Operadora/or de Carpintería y Fabricación de Mobiliario',
            titulo_linea2='Operadora/or de Carpintería y Fabricación de Mobiliario',
        )

        gen = PPTXGenerator()
        out_path = 'output/test_long_trayecto_dashes.pptx'
        gen.generate(data, out_path)

        prs = pptx.Presentation(out_path)
        shape_dict = {shape.shape_id: shape.text_frame.text for shape in prs.slides[0].shapes if shape.has_text_frame}

        self.assertTrue(shape_dict[88].endswith('---'))

        if os.path.exists(out_path):
            os.remove(out_path)

    def test_generate_batch_creates_two_slides_per_title(self):
        titles = [
            TituloData(apellido_nombre='PRUEBA UNO', numero_egresado='1'),
            TituloData(apellido_nombre='PRUEBA DOS', numero_egresado='2'),
        ]
        out_path = 'output/test_batch_generation.pptx'
        PPTXGenerator().generate_batch(titles, out_path)

        try:
            prs = pptx.Presentation(out_path)
            self.assertEqual(len(prs.slides), 4)
            slide_text = [
                shape.text
                for slide in prs.slides
                for shape in slide.shapes
                if shape.has_text_frame
            ]
            self.assertIn('PRUEBA UNO', slide_text)
            self.assertIn('PRUEBA DOS', slide_text)
        finally:
            if os.path.exists(out_path):
                os.remove(out_path)

    def test_generate_batch_copies_template_images(self):
        titles = [TituloData(apellido_nombre='PRUEBA UNO'), TituloData(apellido_nombre='PRUEBA DOS')]
        out_path = 'output/test_batch_images.pptx'
        PPTXGenerator().generate_batch(titles, out_path)

        try:
            prs = pptx.Presentation(out_path)
            self.assertTrue(any(shape.shape_type == 13 for shape in prs.slides[0].shapes))
            self.assertTrue(any(shape.shape_type == 13 for shape in prs.slides[2].shapes))
            self.assertEqual(
                prs.slides[0].shapes[0].image.blob,
                prs.slides[2].shapes[0].image.blob,
            )
        finally:
            if os.path.exists(out_path):
                os.remove(out_path)

if __name__ == '__main__':
    unittest.main()
