#define TIFR3_Bit5 ICF3
#define TIFR3_Bit2 OCF3B
#define TIFR3_Bit1 OCF3A
#define TIFR3_Bit0 TOV3
#define TIFR2_Bit2 OCF2B
#define TIFR2_Bit1 OCF2A
#define TIFR2_Bit0 TOV2
#define TIFR1_Bit5 ICF1
#define TIFR1_Bit2 OCF1B
#define TIFR1_Bit1 OCF1A
#define TIFR1_Bit0 TOV1
#define TIFR0_Bit2 OCF0B
#define TIFR0_Bit1 OCF0A
#define TIFR0_Bit0 TOV0
#define PORTD7 D, 7
#define PORTD6 D, 6
#define PORTD5 D, 5
#define PORTD4 D, 4
#define PORTD3 D, 3
#define PORTD2 D, 2
#define PORTD1 D, 1
#define PORTD0 D, 0
#define DDD7 D, 7
#define DDD6 D, 6
#define DDD5 D, 5
#define DDD4 D, 4
#define DDD3 D, 3
#define DDD2 D, 2
#define DDD1 D, 1
#define DDD0 D, 0
#define PIND7 D, 7
#define PIND6 D, 6
#define PIND5 D, 5
#define PIND4 D, 4
#define PIND3 D, 3
#define PIND2 D, 2
#define PIND1 D, 1
#define PIND0 D, 0
#define PORTC7 C, 7
#define PORTC6 C, 6
#define PORTC5 C, 5
#define PORTC4 C, 4
#define PORTC3 C, 3
#define PORTC2 C, 2
#define PORTC1 C, 1
#define PORTC0 C, 0
#define DDC7 C, 7
#define DDC6 C, 6
#define DDC5 C, 5
#define DDC4 C, 4
#define DDC3 C, 3
#define DDC2 C, 2
#define DDC1 C, 1
#define DDC0 C, 0
#define PINC7 C, 7
#define PINC6 C, 6
#define PINC5 C, 5
#define PINC4 C, 4
#define PINC3 C, 3
#define PINC2 C, 2
#define PINC1 C, 1
#define PINC0 C, 0
#define PORTB7 B, 7
#define PORTB6 B, 6
#define PORTB5 B, 5
#define PORTB4 B, 4
#define PORTB3 B, 3
#define PORTB2 B, 2
#define PORTB1 B, 1
#define PORTB0 B, 0
#define DDB7 B, 7
#define DDB6 B, 6
#define DDB5 B, 5
#define DDB4 B, 4
#define DDB3 B, 3
#define DDB2 B, 2
#define DDB1 B, 1
#define DDB0 B, 0
#define PINB7 B, 7
#define PINB6 B, 6
#define PINB5 B, 5
#define PINB4 B, 4
#define PINB3 B, 3
#define PINB2 B, 2
#define PINB1 B, 1
#define PINB0 B, 0
#define PORTA7 A, 7
#define PORTA6 A, 6
#define PORTA5 A, 5
#define PORTA4 A, 4
#define PORTA3 A, 3
#define PORTA2 A, 2
#define PORTA1 A, 1
#define PORTA0 A, 0
#define DDA7 A, 7
#define DDA6 A, 6
#define DDA5 A, 5
#define DDA4 A, 4
#define DDA3 A, 3
#define DDA2 A, 2
#define DDA1 A, 1
#define DDA0 A, 0
#define PINA7 A, 7
#define PINA6 A, 6
#define PINA5 A, 5
#define PINA4 A, 4
#define PINA3 A, 3
#define PINA2 A, 2
#define PINA1 A, 1
#define PINA0 A, 0

#define HAL_INCLUDES_ADC    1
#define ADC_CHANNEL_REG                     ADC0_MUXPOS
#define ADC_start_conversion()      ADC0_COMMAND |= (1 << ADC_STCONV_bp)
#define ADC_get_status()                    ( !((ADC0_INTFLAGS >> ADC_RESRDY_bp) & 0x01))
#define ADC_get_result()                    ADC0_RES

#define HAL_INCLUDES_DAC    1
#define DAC_enable()                        DAC0_CTRLA |= (1 << DAC_ENABLE_bp)
#define DAC_disable()                       DAC0_CTRLA &= ~(1 << DAC_ENABLE_bp)
#define DAC_output_enable()         DAC0_CTRLA |= (1 << DAC_OUTEN_bp)
#define DAC_output_disable()        DAC0_CTRLA &= ~(1 << DAC_OUTEN_bp)
#define DAC_load(data)                  DAC0_DATA = (data << 6)